$ErrorActionPreference = "Stop"

python -m pip install -r requirements.txt

pyinstaller `
  --name "报关单生成" `
  --windowed `
  --onefile `
  --clean `
  --collect-all webview `
  --add-data "static;static" `
  --add-data "报关单 IMOS 空白模板.xlsx;." `
  --add-data "2026+Daily+Export+List.xlsx;." `
  desktop.py

$PackageDir = Join-Path $PSScriptRoot "dist\报关单生成-客户交付包"
$PackageZip = Join-Path $PSScriptRoot "dist\报关单生成-客户交付包.zip"
$LogicDir = Resolve-Path (Join-Path $PSScriptRoot "..\最新逻辑_请优先看")

if (Test-Path $PackageDir) {
  Remove-Item $PackageDir -Recurse -Force
}
if (Test-Path $PackageZip) {
  Remove-Item $PackageZip -Force
}

New-Item -ItemType Directory -Path $PackageDir | Out-Null
Copy-Item (Join-Path $PSScriptRoot "dist\报关单生成.exe") $PackageDir
Copy-Item (Join-Path $LogicDir "报关单配置关系表.xlsx") $PackageDir
Copy-Item (Join-Path $LogicDir "报关单生成取值逻辑说明.docx") $PackageDir
Copy-Item (Join-Path $LogicDir "报关单配置关系表说明文档.docx") $PackageDir
Copy-Item (Join-Path $LogicDir "README_最新逻辑.md") $PackageDir

Compress-Archive -Path (Join-Path $PackageDir "*") -DestinationPath $PackageZip -Force

Write-Host ""
Write-Host "打包完成：dist\报关单生成.exe"
Write-Host "客户交付包：dist\报关单生成-客户交付包.zip"
