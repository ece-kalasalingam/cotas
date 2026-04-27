/**
 * Cloudflare Worker — CIP generation proxy for COTAS/FOCUS.
 *
 * The EXE sends a compact CO attainment JSON (abbreviated field codes).
 * This Worker holds the Gemini API key as an encrypted environment secret,
 * builds the full prompt, calls gemini-2.5-flash-lite, and returns the CIP text.
 *
 * Required environment secrets (set in Cloudflare dashboard → Workers → Settings → Variables):
 *   GEMINI_API_KEY  — your Google AI Studio key
 *   APP_TOKEN       — a long random string shared with the EXE (kept in EXE config)
 *
 * Rate limiting: configure a Cloudflare Rate Limiting rule in the dashboard
 * (Security → WAF → Rate limiting rules) — 20 req/min per IP is a safe default.
 */

const MAX_REQ_PER_DAY = 200;
const MAX_BODY_BYTES = 64 * 1024; // 64 KB hard cap on incoming payload
const GEMINI_MODEL = "gemini-2.5-flash-lite";
const GEMINI_MAX_OUTPUT_TOKENS = 900;
const GEMINI_TEMPERATURE = 0.4;
const GEMINI_TIMEOUT_MS = 25_000; // abort if Gemini hasn't responded in 25 s
const GEMINI_ENDPOINT =
  `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent`;

// ---------------------------------------------------------------------------
// System instruction sent to Gemini as a fixed context prefix.
// The EXE never sees this — it lives only in the Worker.
// ---------------------------------------------------------------------------
const SYSTEM_INSTRUCTION = `\
You are an expert in Outcome-Based Education (OBE) and NBA accreditation.

You will receive JSON describing Course Outcome (CO) attainment data.

Your task is to generate ONLY the Continuous Improvement Plan (CIP) in a concise format.

## JSON Field Dictionary

course — course identity
  code: course code | sem: semester 
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
  da:      direct attainment score (0-100 scale; weighted average across direct assessments)
  ia:      indirect attainment score (0-100 scale; typically from student survey)
  avg:     combined score = (d_wt x da + i_wt x ia) / 100
  att_pct: % of students who reached tgt_lvl — this is what determines attained/not attained
  st:      "A" = Attained, "NA" = Not Attained
  sf:      shortfall % (how far att_pct fell below tgt_pct; 0 when attained)
  dist:    student level distribution array — [% below L1, % between L1 and L2, % between L2 and L3, % above L3]

## Instructions

For each CO where st = "NA":

1. Identify the ISSUE:
   - Use att_pct, sf, and dist
   - Describe where students are concentrated (L0, L1, etc.)
   - Do NOT assume causes not supported by data

2. Suggest CORRECTIVE ACTION:
   - Use Bloom level (bl) and topics
   - Make actions specific to topics
   - Keep actions practical and must be measurable
   - Keep actions must be implementable (not generic advice)

3. Keep each CO output within 2-3 lines

## Constraints

- Do NOT generate full report sections
- Do NOT infer teaching quality or student behaviour
- Use phrases like "may indicate", "requires review"
- Do NOT invent data
- Keep total output under 300 words

## Output Format

CO ID: <id>
Issue: <issue>
Action: <action>

Repeat for all NOT attained COs. \
Do not state unsupported causes as facts.\
Do not invent numbers or outcomes not present in the JSON.`;

// ---------------------------------------------------------------------------
// Main entry point
// ---------------------------------------------------------------------------
export default {
  async fetch(request, env) {
    if (env.DISABLE_GEMINI === "1") {
      return Response.json({ error: "CIP disabled" }, { status: 503 });
    }
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

    // 4. Max request rate per day check (simple in-memory counter — resets on Worker restart).
    const key = "count:" + new Date().toISOString().slice(0,10); // YYYY-MM-DD
    const current = Number(await env.KV.get(key) || "0");

    if (current >= MAX_REQ_PER_DAY) {
      return new Response(JSON.stringify({ error: "Daily limit reached" }), { status: 429 });
    }

    await env.KV.put(key, String(current + 1), { expirationTtl: 86400 });


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
      if (new TextEncoder().encode(raw).length > MAX_BODY_BYTES) {
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
    if (!env.KV) {
      console.warn("KV not available, skipping cache");
    }
    const cacheKey = env.KV ? await sha256(JSON.stringify(payload, Object.keys(payload).sort())) : null;
    if (env.KV && cacheKey) {
      const cached = await env.KV.get(cacheKey);
      if (cached) return Response.json({ report_text: cached, cached: true });
    }
    try {
      console.log("Request:", payload.course?.code);
      const cipText = await callGemini(payload, env.GEMINI_API_KEY);
      if (env.KV && cacheKey) {
        await env.KV.put(cacheKey, cipText, { expirationTtl: 180 * 86400 }); // cache for 180 days
      }
      return json(200, { report_text: cipText, cached: false });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      return json(502, { error: "Gemini call failed", detail: msg.slice(0,200) });
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

async function sha256(text) {
  const data = new TextEncoder().encode(text);
  const hash = await crypto.subtle.digest("SHA-256", data);
  return [...new Uint8Array(hash)]
    .map(b => b.toString(16).padStart(2, "0"))
    .join("");
}
