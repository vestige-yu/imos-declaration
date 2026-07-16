# 报表合并

这是“报表处理工具”中的第二个业务模块。正式客户版本会把它与“报关单生成”打包进同一个 Windows EXE；本目录仍可独立运行和测试，便于维护合并引擎。

当前能力：

- 一次合并 1～30 份 `.xls` / `.xlsx` 源报表。
- 客户通过 `.xlsx` 上传并定义输出模板。
- 客户通过 `.xlsx` 上传并定义 QAD PN 匹配规则。
- 自动识别源报表明细表头、输出模板表头和规则工作表。
- 应用客户确认的料号命名格式规则，并对非原文精确匹配保留审计提示。
- 生成前预览全部汇总、匹配状态和异常；黄色或红色异常必须人工确认。
- 生成 `.xlsx`，同时返回条数、数量、金额、分单汇总和异常警告。
- 保留本机历史记录，包括源报表、当时使用的模板和规则、预览 JSON 与最终结果。
- 不区分管理员和普通用户。

## 目录结构

- `src/`：应用程序代码。
- `tests/`：自动化测试和测试辅助文件。
- `samples/`：输入报表样例；不要放客户正式数据或敏感数据。
- `outputs/`：本地生成结果，不提交到 Git。
- `docs/`：需求、字段映射和合并逻辑说明。
- `static/`：本地浏览器页面。
- `app.py`：本地 HTTP 服务和历史记录。
- `desktop.py`：Windows 桌面入口。

## 本地运行

先安装依赖：

```bash
python3 -m pip install -r requirements.txt
```

在本目录执行：

```bash
python3 app.py
```

然后打开：

```text
http://127.0.0.1:8000/
```

需要自动打开默认浏览器时：

```bash
SURI_OPEN_BROWSER=1 python3 app.py
```

仍可使用 `merge_reports.py` 直接执行命令行合并，参数说明运行 `python3 merge_reports.py --help`。

## 数据保存位置

Windows 默认位置：

```text
%LOCALAPPDATA%\SuriWorkReportMerge\
```

主要内容：

```text
app.db
config.json
templates\current.xlsx
templates\versions\...
rules\current.xlsx
rules\versions\...
history\<记录ID>\sources\...
history\<记录ID>\template.xlsx
history\<记录ID>\rules.xlsx
history\<记录ID>\preview.json
history\<记录ID>\output.xlsx
```

可通过环境变量临时指定：

```bat
set SURI_REPORT_MERGE_DATA_DIR=D:\SuriWorkReportMergeData
```

历史记录包含客户上传文件，删除历史记录会同时删除对应本地文件。

## 运行测试

```bash
python3 -m unittest discover -s tests -v
```

## 独立调试打包

在 Windows PowerShell 中进入本目录：

```powershell
.\build_windows.ps1
```

该脚本只用于单独调试本模块，生成：

```text
dist\报表合并.exe
dist\报表合并-客户交付包.zip
```

正式客户交付请使用 `.github/workflows/build-local-desktop-exe.yml`，由统一入口生成 `报表处理工具.exe` 和客户交付包。

## 预览和异常

- 精确匹配：正常显示。
- 命名规则匹配：黄色显示，并展示实际使用的规则料号和匹配原因。
- 完全未匹配或归一后对应多条冲突规则：红色显示，商品名称、商品编号和品牌留空。
- 模板含未支持字段：黄色提醒；归一后的规则字段冲突属于红色异常。
- 只要存在任何黄色或红色提醒，用户必须勾选确认后才能生成。

模板允许客户调整已支持字段的顺序、表头名称和 Excel 样式。不支持任意业务字段、任意公式或多级表头。

## 料号命名规则

代码按 `料号命名格式确认表.xlsx` 中的客户确认口径执行：

- 视为同一料号：分隔符不同、首尾零不同、版本后缀不同、全角/空白差异、Excel 数字 `.0` 差异。
- 视为不同料号：客户或品牌前缀不同、字母与数字相似（如 O/0、I/1）、无明确映射的新旧料号。
- 原文精确匹配优先。
- 非原文精确匹配只有在归一后的候选规则字段一致时才自动使用，并标黄。
- 如果归一后候选规则的商品名称、商品编号或品牌不一致，则标红，不自动任选。

## 开发约定

1. 报表合并的业务引擎、依赖和默认配置都放在本目录内。
2. 统一桌面应用通过 `local_desktop_app/merge_bridge.py` 加载本模块；合并业务逻辑仍保持独立，不与报关单逻辑混写。
3. 合并规则发生变化时，同步更新 `docs/合并逻辑说明.md`。
4. 测试样例和预期结果应成对保存，便于回归验证。
