param(
    [string]$AppVersion = "1.0.5",
    [switch]$SkipContributorsRefresh,
    [switch]$SkipBuild
)

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
    }

    if (-not $buildSuccess) {
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

$isccArgs = @(
    "/DAppVersion=$AppVersion"
    "installer\focus.iss"
)

& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" $isccArgs
exit $LASTEXITCODE
