# Windows 本地版 IMOS 报关单生成工具

这是一套和线上网站隔离的本地桌面版文件。原网站根目录的 `app.py` 和 `static/` 不需要改动。

## 本地运行

```bash
python app.py
```

打开：

```text
http://127.0.0.1:8000/
```

桌面窗口运行：

```bash
python desktop.py
```

## 数据保存位置

Windows 默认保存到：

```text
%LOCALAPPDATA%\SuriWorkDeclaration\
```

目录结构：

```text
app.db
templates\template.xlsx
rules\rules.xlsx
history\<记录ID>\invoice.xls
history\<记录ID>\packing.xls
history\<记录ID>\preview.json
history\<记录ID>\output.xlsx
outputs\
```

如果需要临时指定数据目录，可以设置：

```bash
set SURI_DATA_DIR=D:\SuriWorkDeclarationData
```

## Windows 打包

在 Windows 电脑进入这个目录后运行：

```powershell
.\build_windows.ps1
```

完成后生成：

```text
dist\IMOS报关单生成.exe
```

## 安全边界

- 桌面版只监听 `127.0.0.1`，不会暴露到公网。
- Invoice、Packing list、预览数据、生成的报关单都会保存在本机历史记录里。
- 不做账号登录，数据保护依赖 Windows 当前用户账号。
