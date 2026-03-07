if (-not $env:WB_PWD) {
    Write-Error "WB_PWD environment variable is not set. Aborting build."
    exit 1
}

$args = @(
    "/DWorkbookPassword=$env:WB_PWD"
    "/DAppVersion=1.0.0"
    "installer\focus.iss"
)

& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" $args