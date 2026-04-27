param(
    [string]$AppVersion = "",
    [switch]$SkipContributorsRefresh,
    [switch]$SkipBuild
)

function Resolve-AppVersion {
    param(
        [string]$RequestedVersion
    )

    if (-not [string]::IsNullOrWhiteSpace($RequestedVersion)) {
        return $RequestedVersion.Trim()
    }

    $resolved = ""
    $conda = Get-Command conda -ErrorAction SilentlyContinue
    if ($conda) {
        $resolved = (conda run -n obe python -c "from common.constants import SYSTEM_VERSION; print(SYSTEM_VERSION)" 2>$null | Select-Object -Last 1).Trim()
    }

    if ([string]::IsNullOrWhiteSpace($resolved)) {
        $python = Get-Command python -ErrorAction SilentlyContinue
        if ($python) {
            $resolved = (python -c "from common.constants import SYSTEM_VERSION; print(SYSTEM_VERSION)" 2>$null | Select-Object -Last 1).Trim()
        }
    }

    if ([string]::IsNullOrWhiteSpace($resolved)) {
        Write-Error "Could not resolve SYSTEM_VERSION from common/constants.py. Pass -AppVersion explicitly."
        exit 1
    }

    return $resolved
}

function Resolve-AppFileVersion {
    param(
        [string]$SemanticVersion
    )

    $parts = ($SemanticVersion -split '\.')
    $numericParts = @()

    foreach ($part in $parts) {
        if ($part -match '^(\d+)') {
            $numericParts += [int]$Matches[1]
        } else {
            $numericParts += 0
        }
    }

    while ($numericParts.Count -lt 4) {
        $numericParts += 0
    }

    if ($numericParts.Count -gt 4) {
        $numericParts = $numericParts[0..3]
    }

    return ($numericParts -join '.')
}

$ResolvedAppVersion = Resolve-AppVersion -RequestedVersion $AppVersion
$ResolvedAppFileVersion = Resolve-AppFileVersion -SemanticVersion $ResolvedAppVersion
Write-Host "Using installer AppVersion: $ResolvedAppVersion"
Write-Host "Using installer AppFileVersion: $ResolvedAppFileVersion"

if (-not $SkipContributorsRefresh) {
    Write-Host "Refreshing contributors list for About module..."
    $contributorsRefreshed = $false

    $conda = Get-Command conda -ErrorAction SilentlyContinue
    if ($conda) {
        conda run -n obe python common/get_contributors.py `
            --owner ece-kalasalingam `
            --repo cotas `
            --output assets/about_contributors.txt
        if ($LASTEXITCODE -eq 0) {
            $contributorsRefreshed = $true
        }
    }

    if (-not $contributorsRefreshed) {
        $python = Get-Command python -ErrorAction SilentlyContinue
        if ($python) {
            python common/get_contributors.py `
                --owner ece-kalasalingam `
                --repo cotas `
                --output assets/about_contributors.txt
            if ($LASTEXITCODE -eq 0) {
                $contributorsRefreshed = $true
            }
        }
    }

    if (-not $contributorsRefreshed) {
        Write-Warning "Could not refresh contributors list. Installer build will continue with existing contributors file."
    }
}

if (-not $SkipBuild) {
    function Invoke-ReleaseGateCommand {
        param(
            [string]$Name,
            [string[]]$CondaArgs,
            [string[]]$PythonArgs
        )

        Write-Host "Running $Name..."
        $conda = Get-Command conda -ErrorAction SilentlyContinue
        if ($conda) {
            conda run -n obe @CondaArgs
            if ($LASTEXITCODE -eq 0) {
                return $true
            }
        } else {
            $python = Get-Command python -ErrorAction SilentlyContinue
            if ($python) {
                python @PythonArgs
                if ($LASTEXITCODE -eq 0) {
                    return $true
                }
            }
        }

        Write-Error "$Name failed. Aborting installer packaging."
        return $false
    }

    Write-Host "Running quality gate..."
    $qualityGatePassed = $false
    $conda = Get-Command conda -ErrorAction SilentlyContinue
    if ($conda) {
        conda run -n obe python scripts/quality_gate.py
        if ($LASTEXITCODE -eq 0) {
            $qualityGatePassed = $true
        }
    } else {
        $python = Get-Command python -ErrorAction SilentlyContinue
        if ($python) {
            python scripts/quality_gate.py
            if ($LASTEXITCODE -eq 0) {
                $qualityGatePassed = $true
            }
        }
    }

    if (-not $qualityGatePassed) {
        Write-Error "Quality gate failed. Aborting installer packaging."
        exit 1
    }

    if (-not (Invoke-ReleaseGateCommand -Name "Ruff check" -CondaArgs @("python", "-m", "ruff", "check", ".") -PythonArgs @("-m", "ruff", "check", "."))) {
        exit 1
    }
    if (-not (Invoke-ReleaseGateCommand -Name "isort check" -CondaArgs @("python", "-m", "isort", "--check-only", "--diff", ".") -PythonArgs @("-m", "isort", "--check-only", "--diff", "."))) {
        exit 1
    }
    if (-not (Invoke-ReleaseGateCommand -Name "Pyright" -CondaArgs @("python", "-m", "pyright") -PythonArgs @("-m", "pyright"))) {
        exit 1
    }
    if (-not (Invoke-ReleaseGateCommand -Name "Bandit" -CondaArgs @("python", "-m", "bandit", "-q", "-c", ".bandit.yaml", "-r", "common", "modules", "services") -PythonArgs @("-m", "bandit", "-q", "-c", ".bandit.yaml", "-r", "common", "modules", "services"))) {
        exit 1
    }
    if (-not (Invoke-ReleaseGateCommand -Name "pip-audit" -CondaArgs @("python", "-m", "pip_audit", "--cache-dir", ".pip_audit_cache", "--ignore-vuln", "GHSA-58qw-9mgm-455v") -PythonArgs @("-m", "pip_audit", "--cache-dir", ".pip_audit_cache", "--ignore-vuln", "GHSA-58qw-9mgm-455v"))) {
        exit 1
    }

    Write-Host "Generating version file and building with PyInstaller..."
    $buildSuccess = $false

    $conda = Get-Command conda -ErrorAction SilentlyContinue
    if ($conda) {
        conda run -n obe python scripts/generate_version_file.py
        if ($LASTEXITCODE -eq 0) {
            conda run -n obe python -m PyInstaller --noconfirm --clean FOCUS.spec
            if ($LASTEXITCODE -eq 0) {
                $buildSuccess = $true
            }
        }
    } else {
        $python = Get-Command python -ErrorAction SilentlyContinue
        if ($python) {
            python scripts/generate_version_file.py
            if ($LASTEXITCODE -eq 0) {
                python -m PyInstaller --noconfirm --clean FOCUS.spec
                if ($LASTEXITCODE -eq 0) {
                    $buildSuccess = $true
                }
            }
        }
    }

    if (-not $buildSuccess) {
        Write-Error "PyInstaller build failed. Aborting installer packaging."
        exit 1
    }
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$cipConfigSource = Join-Path $repoRoot "cip_config.json"
$distRoot = Join-Path $repoRoot "dist\focus"
$cipConfigTarget = Join-Path $distRoot "cip_config.json"

if (-not (Test-Path -LiteralPath $cipConfigSource)) {
    Write-Error "Missing required cip_config.json at repo root. Aborting installer packaging."
    exit 1
}

if (-not (Test-Path -LiteralPath $distRoot)) {
    Write-Error "Missing PyInstaller output directory '$distRoot'. Aborting installer packaging."
    exit 1
}

Copy-Item -LiteralPath $cipConfigSource -Destination $cipConfigTarget -Force
Write-Host "Staged cip_config.json to $cipConfigTarget"

$portableArchivePath = Join-Path $repoRoot "dist\portableexefile.zip"
if (Test-Path -LiteralPath $portableArchivePath) {
    Remove-Item -LiteralPath $portableArchivePath -Force
}

$portableArchiveSource = Join-Path $distRoot "*"
Compress-Archive -Path $portableArchiveSource -DestinationPath $portableArchivePath -CompressionLevel Optimal
Write-Host "Built portable archive: $portableArchivePath"

$isccArgs = @(
    "/DAppVersion=$ResolvedAppVersion"
    "/DAppFileVersion=$ResolvedAppFileVersion"
    "installer\focus.iss"
)

& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" $isccArgs
exit $LASTEXITCODE
