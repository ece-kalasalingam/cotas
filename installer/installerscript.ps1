$env:WB_PWD = 'ece@KLU131984'

$args = @(
    "/DWorkbookPassword=$env:WB_PWD"
    "/DAppVersion=1.0.0"
    "installer\focus.iss"
)

& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" $args