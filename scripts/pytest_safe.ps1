param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$PytestArgs
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$safeRoot = Join-Path $repoRoot ".pytest_temp_cotas"
$baseTemp = Join-Path $safeRoot "basetemp"
$cacheDir = Join-Path $repoRoot ".pytest_cache_cotas"

New-Item -ItemType Directory -Force -Path $safeRoot | Out-Null
New-Item -ItemType Directory -Force -Path $baseTemp | Out-Null
New-Item -ItemType Directory -Force -Path $cacheDir | Out-Null

# Keep temp writes in a repo-owned location for this process only.
$env:TEMP = $safeRoot
$env:TMP = $safeRoot

$args = @("-m", "pytest", "--basetemp", $baseTemp)
if ($PytestArgs) {
    $args += $PytestArgs
}

conda run -n obe python @args
