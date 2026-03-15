param(
    [string]$AppVersion = "1.0.0"
)

$args = @(
    "/DAppVersion=$AppVersion"
    "installer\focus.iss"
)

& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" $args
