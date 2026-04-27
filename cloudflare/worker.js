/**
 * Cloudflare Worker — CIP generation proxy for COTAS/FOCUS.
 *
 * The EXE sends a compact CO attainment JSON (abbreviated field codes).
 * This Worker holds the Gemini API key as an encrypted environment secret,
 * builds the full prompt, calls gemini-2.0-flash, and returns the CIP text.
 *
 * Required environment secrets (set in Cloudflare dashboard → Workers → Settings → Variables):
 *   GEMINI_API_KEY  — your Google AI Studio key
 *   APP_TOKEN       — a long random string shared with the EXE (kept in EXE config)
 *
 * Rate limiting: configure a Cloudflare Rate Limiting rule in the dashboard
 * (Security → WAF → Rate limiting rules) — 20 req/min per IP is a safe default.
 */

const MAX_BODY_BYTES = 64 * 1024; // 64 KB hard cap on incoming payload
const GEMINI_MODEL = "gemini-2.0-flash";
const GEMINI_MAX_OUTPUT_TOKENS = 2048;
const GEMINI_TEMPERATURE = 0.4;
const GEMINI_TIMEOUT_MS = 25_000; // abort if Gemini hasn't responded in 25 s
const GEMINI_ENDPOINT =
  `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent`;

// ---------------------------------------------------------------------------
// System instruction sent to Gemini as a fixed context prefix.
// The EXE never sees this — it lives only in the Worker.
// ---------------------------------------------------------------------------
const SYSTEM_INSTRUCTION = `\
You are an expert in Outcome-Based Education (OBE) and NBA (National Board of \
Accreditation) accreditation for Indian engineering colleges. You will receive a \
JSON object describing Course Outcome (CO) attainment data for one course-section. \
Write a formal Continuous Improvement Plan (CIP) suitable for direct inclusion in \
the official NBA accreditation Word report.

## JSON Field Dictionary

course — course identity
  code: course code | sem: semester | sec: section
  ay: academic year | students: total enrolled students

policy — attainment policy
  tgt_pct: target attainment percentage (e.g. 80 means 80 % of students must reach tgt_lvl)
  tgt_lvl: target level ("L1"/"L2"/"L3") mapped to the corresponding threshold score
  thresh: [L1_score, L2_score, L3_score] — score thresholds that define the three levels
  d_wt: direct assessment weight %  |  i_wt: indirect assessment weight %

summary
  total: total number of COs | att: COs attained | not_att: COs not attained

assessments — list of assessment components used in this course
  name: component name | wt: maximum marks / weightage
  d: true = direct assessment, false = indirect (e.g. student survey)

cos — one entry per Course Outcome
  id:      CO identifier (e.g. "ECE001.3")
  desc:    CO description — what the student must be able to do
  bl:      Bloom's taxonomy cognitive level
             1=Remember  2=Understand  3=Apply  4=Analyze  5=Evaluate  6=Create
  topics:  key topics / experiments / projects covered under this CO
  da:      direct attainment score (0–100 scale; weighted average across direct assessments)
  ia:      indirect attainment score (0–100 scale; typically from student survey)
  avg:     combined score = (d_wt × da + i_wt × ia) / 100
  att_pct: % of students who reached tgt_lvl — this is what determines attained/not attained
  st:      "A" = Attained, "NA" = Not Attained
  sf:      shortfall % (how far att_pct fell below tgt_pct; 0 when attained)
  dist:    student level distribution array — [% below L1, % at L1, % at L2, % at L3]

## Output Format

Write the CIP as structured prose with the following five sections. Use the exact \
headings shown. Do not include JSON, code blocks, or markdown in your response — \
plain text with numbered headings only.

1. Overall Attainment Summary
   Two to three sentences on the overall attainment status, how many COs were attained \
vs not attained, and the general performance pattern.

2. CO-wise Analysis
   For each not-attained CO write one focused paragraph: state the CO id and description, \
quantify the shortfall, interpret the dist array to describe where students cluster, \
identify the likely root cause (content gap, assessment difficulty, Bloom's level \
mismatch, etc.), and propose one or two targeted corrective actions specific to that CO.

3. Assessment Strategy
   Recommendations on assessment design — question difficulty calibration, marks \
distribution across Bloom's levels, mapping of questions to COs, and whether the \
current mix of direct/indirect assessments is adequate.

4. Teaching-Learning Improvements
   Specific pedagogical actions: active learning techniques, additional practice \
resources, remedial interventions, lab/project enhancements, or feedback mechanisms.

5. Action Plan for Next Cycle
   A numbered list of 6–8 concrete, measurable actions with the responsible party \
(faculty / department / board of studies) noted for each item.

Keep the language formal, precise, and free of speculation beyond what the data supports. \
Do not invent numbers or outcomes not present in the JSON.`;

// ---------------------------------------------------------------------------
// Main entry point
// ---------------------------------------------------------------------------
export default {
  async fetch(request, env) {
    // 1. Only POST is accepted.
    if (request.method !== "POST") {
      return json(405, { error: "Method not allowed" });
    }

    // 2. Block browser-origin requests — prevents direct browser calls.
    if (request.headers.get("Origin")) {
      return json(403, { error: "Browser origin not permitted" });
    }

    // 3. Validate the shared app token.
    const token = request.headers.get("X-App-Token") || "";
    if (!env.APP_TOKEN || token !== env.APP_TOKEN) {
      return json(401, { error: "Unauthorized" });
    }

    // 4. Require JSON content type.
    const ct = request.headers.get("Content-Type") || "";
    if (!ct.includes("application/json")) {
      return json(415, { error: "Content-Type must be application/json" });
    }

    // 5. Enforce payload size limit (Content-Length header fast-path).
    const clHeader = parseInt(request.headers.get("Content-Length") || "0", 10);
    if (clHeader > MAX_BODY_BYTES) {
      return json(413, { error: "Payload too large" });
    }

    // 6. Read and parse the body.
    let payload;
    try {
      const raw = await request.text();
      if (raw.length > MAX_BODY_BYTES) {
        return json(413, { error: "Payload too large" });
      }
      payload = JSON.parse(raw);
    } catch {
      return json(400, { error: "Invalid JSON body" });
    }

    // 7. Minimal structure check — must have course and cos array.
    if (
      !payload.course ||
      !Array.isArray(payload.cos) ||
      payload.cos.length === 0
    ) {
      return json(400, { error: "Missing required fields: course, cos" });
    }

    // 8. Guard: Gemini key must be configured as a Worker secret.
    if (!env.GEMINI_API_KEY) {
      return json(500, { error: "Gemini API key not configured" });
    }

    // 9. Call Gemini and return the CIP text.
    try {
      const cipText = await callGemini(payload, env.GEMINI_API_KEY);
      return json(200, { report_text: cipText });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      return json(502, { error: "Gemini call failed", detail: msg });
    }
  },
};

// ---------------------------------------------------------------------------
// Gemini API call
// ---------------------------------------------------------------------------
async function callGemini(payload, apiKey) {
  const userMessage =
    "Here is the CO attainment data for the course. Write the CIP.\n\n" +
    JSON.stringify(payload);

  const body = {
    system_instruction: { parts: [{ text: SYSTEM_INSTRUCTION }] },
    contents: [{ role: "user", parts: [{ text: userMessage }] }],
    generationConfig: {
      temperature: GEMINI_TEMPERATURE,
      maxOutputTokens: GEMINI_MAX_OUTPUT_TOKENS,
    },
  };

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), GEMINI_TIMEOUT_MS);

  let response;
  try {
    response = await fetch(`${GEMINI_ENDPOINT}?key=${apiKey}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
      signal: controller.signal,
    });
  } catch (err) {
    if (err.name === "AbortError") {
      throw new Error(`Gemini request timed out after ${GEMINI_TIMEOUT_MS} ms`);
    }
    throw err;
  } finally {
    clearTimeout(timer);
  }

  if (!response.ok) {
    const errText = await response.text().catch(() => "(no body)");
    throw new Error(`Gemini ${response.status}: ${errText.slice(0, 200)}`);
  }

  const data = await response.json();
  const text =
    data?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) {
    throw new Error("Gemini returned no text content");
  }
  return text.trim();
}

// ---------------------------------------------------------------------------
// Helper
// ---------------------------------------------------------------------------
function json(status, body) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}
