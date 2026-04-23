param(
    [string]$AppVersion = "1.0.5",
    [switch]$SkipContributorsRefresh
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

$args = @(
    "/DAppVersion=$AppVersion"
    "installer\focus.iss"
)

& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" $args
exit $LASTEXITCODE
