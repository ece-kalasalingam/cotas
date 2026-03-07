param(
    [string]$WorkbookPassword = $env:WB_PWD,
    [string]$AppVersion = "1.0.0"
)

if ([string]::IsNullOrWhiteSpace($WorkbookPassword)) {
    throw "Workbook password is required. Set WB_PWD in environment or pass -WorkbookPassword."
}

$args = @(
    "/DWorkbookPassword=$WorkbookPassword"
    "/DAppVersion=$AppVersion"
    "installer\focus.iss"
)

& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" $args
