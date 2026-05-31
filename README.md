# IMOS 报关单生成网站

这是一个零第三方依赖的轻量网站，用于上传 `Invoice` 和 `Packing list`，预览解析结果，并生成可下载的报关单 `.xlsx`。

## 本地运行

```bash
python3 app.py
```

打开：

- 外部上传页面：`http://localhost:8000/u/imos-demo`
- 管理页面：`http://localhost:8000/admin/admin-demo`

生产环境建议用环境变量替换默认访问 token：

```bash
PUBLIC_TOKEN=一段随机外部访问码 ADMIN_TOKEN=一段随机管理员访问码 PORT=8000 python3 app.py
```

## 文件来源

默认读取当前目录下的：

- `报关单 IMOS 空白模板.xlsx`
- `2026+Daily+Export+List.xlsx`

也可以在管理页面上传新的模板和规则表。上传后会保存到 `storage/`，后续生成优先使用上传版本。

## 接口

- `POST /api/parse`：上传 `invoice` 和 `packing`，返回预览数据和异常提示。
- `POST /api/generate`：传入 `sessionId` 或预览数据，生成并返回下载地址。
- `POST /api/admin/rules`：上传 `template` 或 `rules`，替换后台使用版本。

## 部署提示

初版可以部署到支持 Python HTTP 服务的平台，例如 Render、Railway 或自有服务器。上传文件会保存在临时目录中，生成文件保存在 `outputs/` 并通过一次下载链接返回。

## Render 上线

1. 把这些文件上传到一个 GitHub 仓库：
   - `app.py`
   - `static/`
   - `requirements.txt`
   - `render.yaml`
   - `报关单 IMOS 空白模板.xlsx`
   - `2026+Daily+Export+List.xlsx`

2. 在 Render 新建 Blueprint 或 Web Service，选择这个 GitHub 仓库。

3. 设置环境变量：
   - `PUBLIC_TOKEN`：外部用户访问码，例如一段 20 位以上随机字符串。
   - `ADMIN_TOKEN`：管理员访问码，必须和外部访问码不同。

4. 部署完成后访问：
   - 外部页面：`https://你的服务名.onrender.com/u/PUBLIC_TOKEN`
   - 管理页面：`https://你的服务名.onrender.com/admin/ADMIN_TOKEN`

注意：如果使用普通无持久磁盘部署，管理页面上传的新模板/规则表在重新部署后可能丢失。正式长期使用时，建议把最新模板和规则表直接提交到仓库，或配置持久磁盘/对象存储。
