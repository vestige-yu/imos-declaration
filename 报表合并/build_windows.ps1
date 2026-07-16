$ErrorActionPreference = "Stop"

python -m pip install -r requirements.txt

pyinstaller `
  --name "报表合并" `
  --windowed `
  --onefile `
  --clean `
  --paths "src" `
  --collect-submodules "report_merge" `
  --add-data "src;src" `
  --add-data "static;static" `
  --add-data "2026 Daily Export List模板.xlsx;." `
  --add-data "List模板.xlsx;." `
  desktop.py

$PackageDir = Join-Path $PSScriptRoot "dist\报表合并-客户交付包"
$PackageZip = Join-Path $PSScriptRoot "dist\报表合并-客户交付包.zip"

if (Test-Path $PackageDir) {
  Remove-Item $PackageDir -Recurse -Force
}
if (Test-Path $PackageZip) {
  Remove-Item $PackageZip -Force
}

New-Item -ItemType Directory -Path $PackageDir | Out-Null
Copy-Item (Join-Path $PSScriptRoot "dist\报表合并.exe") $PackageDir
Copy-Item (Join-Path $PSScriptRoot "2026 Daily Export List模板.xlsx") $PackageDir
Copy-Item (Join-Path $PSScriptRoot "List模板.xlsx") $PackageDir
Copy-Item (Join-Path $PSScriptRoot "README.md") $PackageDir
Copy-Item (Join-Path $PSScriptRoot "docs\合并逻辑说明.md") $PackageDir

Compress-Archive `
  -Path (Join-Path $PackageDir "*") `
  -DestinationPath $PackageZip `
  -Force

Write-Host ""
Write-Host "打包完成：dist\报表合并.exe"
Write-Host "客户交付包：dist\报表合并-客户交付包.zip"
