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

## 阿里云上线

当前生产环境部署在阿里云轻量应用服务器，使用 Ubuntu + Python + systemd + Nginx。

服务器上代码目录：

```bash
/opt/imos-declaration
```

后台服务：

```bash
imos-declaration.service
```

常用维护命令：

```bash
sudo systemctl status imos-declaration --no-pager
sudo systemctl restart imos-declaration
sudo journalctl -u imos-declaration -n 100 --no-pager
```

生产环境变量保存在服务器：

```bash
/etc/imos-declaration/app.env
```

需要配置：

- `PUBLIC_TOKEN`：外部用户访问码
- `ADMIN_TOKEN`：管理员访问码
- `PORT`：应用监听端口，当前为 `8000`
- `MAX_UPLOAD_BYTES`：最大上传文件大小

Nginx 负责把公网 `80` 端口转发到本机 Python 服务：

```text
公网访问 -> Nginx :80 -> 127.0.0.1:8000
```

更新服务器代码时：

```bash
cd /opt/imos-declaration
git pull
sudo systemctl restart imos-declaration
```

注意：管理页面上传的新模板/规则表会保存在服务器 `storage/` 目录。长期正式使用时，建议把最新模板和规则表提交到 GitHub，或者定期备份服务器上的 `storage/`。
