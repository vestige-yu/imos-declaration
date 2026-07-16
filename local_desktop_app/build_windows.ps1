$ErrorActionPreference = "Stop"

$RepoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$LogicDir = Resolve-Path (Join-Path $RepoRoot "最新逻辑_请优先看")
$ConfigFile = Resolve-Path (Join-Path $LogicDir "报关单配置关系表.xlsx")
$DeclarationTemplateFile = Resolve-Path (Join-Path $PSScriptRoot "报关单 IMOS 空白模板.xlsx")
$MarkedTemplateFile = Resolve-Path (Join-Path $RepoRoot "报关单模板-标记版.xlsx")
$MarkedTemplateGuide = Resolve-Path (Join-Path $RepoRoot "模板标记维护说明.md")
$MergeDir = Resolve-Path (Join-Path $RepoRoot "报表合并")
$MergeSrcDir = Resolve-Path (Join-Path $MergeDir "src")
$MergeAppFile = Resolve-Path (Join-Path $MergeDir "app.py")
$MergeTemplateFile = Resolve-Path (Join-Path $MergeDir "2026 Daily Export List模板.xlsx")
$MergeRulesFile = Resolve-Path (Join-Path $MergeDir "List模板.xlsx")
$MergeLogicGuide = Resolve-Path (Join-Path $MergeDir "docs\合并逻辑说明.md")

python -m pip install -r requirements.txt

pyinstaller `
  --name "报表处理工具" `
  --windowed `
  --onefile `
  --clean `
  --paths "$MergeSrcDir" `
  --collect-submodules "report_merge" `
  --hidden-import "openpyxl" `
  --add-data "static;static" `
  --add-data "报关单 IMOS 空白模板.xlsx;." `
  --add-data "$ConfigFile;." `
  --add-data "2026+Daily+Export+List.xlsx;." `
  --add-data "$MergeAppFile;merge_module" `
  --add-data "$MergeSrcDir;report_merge_src" `
  --add-data "$MergeTemplateFile;." `
  --add-data "$MergeRulesFile;." `
  desktop.py

$PackageDir = Join-Path $PSScriptRoot "dist\报表处理工具-客户交付包"
$PackageZip = Join-Path $PSScriptRoot "dist\报表处理工具-客户交付包.zip"

if (Test-Path $PackageDir) {
  Remove-Item $PackageDir -Recurse -Force
}
if (Test-Path $PackageZip) {
  Remove-Item $PackageZip -Force
}

New-Item -ItemType Directory -Path $PackageDir | Out-Null
New-Item -ItemType Directory -Path (Join-Path $PackageDir "报关单配置资料") | Out-Null
New-Item -ItemType Directory -Path (Join-Path $PackageDir "报表合并配置资料") | Out-Null

Copy-Item (Join-Path $PSScriptRoot "dist\报表处理工具.exe") $PackageDir
Copy-Item (Join-Path $PSScriptRoot "README.md") (Join-Path $PackageDir "README_客户使用说明.md")

$DeclarationPackageDir = Join-Path $PackageDir "报关单配置资料"
Copy-Item (Join-Path $LogicDir "报关单配置关系表.xlsx") $DeclarationPackageDir
Copy-Item (Join-Path $LogicDir "报关单生成取值逻辑说明.docx") $DeclarationPackageDir
Copy-Item (Join-Path $LogicDir "报关单配置关系表说明文档.docx") $DeclarationPackageDir
Copy-Item (Join-Path $LogicDir "README_最新逻辑.md") $DeclarationPackageDir
Copy-Item $DeclarationTemplateFile (Join-Path $DeclarationPackageDir "报关单生成模板.xlsx")
Copy-Item $MarkedTemplateFile (Join-Path $DeclarationPackageDir "报关单模板-标记版（维护参考）.xlsx")
Copy-Item $MarkedTemplateGuide $DeclarationPackageDir

$MergePackageDir = Join-Path $PackageDir "报表合并配置资料"
Copy-Item $MergeTemplateFile (Join-Path $MergePackageDir "报表合并生成模板.xlsx")
Copy-Item $MergeRulesFile (Join-Path $MergePackageDir "报表合并规则.xlsx")
Copy-Item $MergeLogicGuide $MergePackageDir

Compress-Archive -Path (Join-Path $PackageDir "*") -DestinationPath $PackageZip -Force

Write-Host ""
Write-Host "打包完成：dist\报表处理工具.exe"
Write-Host "客户交付包：dist\报表处理工具-客户交付包.zip"
