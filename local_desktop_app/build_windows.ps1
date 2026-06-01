$ErrorActionPreference = "Stop"

python -m pip install -r requirements.txt

pyinstaller `
  --name "IMOS报关单生成" `
  --windowed `
  --onefile `
  --add-data "static;static" `
  --add-data "报关单 IMOS 空白模板.xlsx;." `
  --add-data "2026+Daily+Export+List.xlsx;." `
  desktop.py

Write-Host ""
Write-Host "打包完成：dist\IMOS报关单生成.exe"
