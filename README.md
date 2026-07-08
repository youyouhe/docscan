# DocScan — Word 文档转 PDF + Markdown API

将 `.docx` 文档精确转换为 **PDF**（保留排版/字体/表格）并提取逐页 **Markdown**（含表格语义）。

## 快速开始

```bash
./start.sh          # 首次自动拉取 ONLYOFFICE (~3GB) + 启动服务
./start.sh 8080     # 指定端口
./stop.sh           # 停止
./restart.sh        # 重启
```

首次运行会自动：
1. `docker compose up -d` 拉取并启动 ONLYOFFICE 容器
2. 禁用 JWT、安装中文字体
3. 启动 DocScan API

启动后访问：
- **API 文档**: http://localhost:8800/api/docs (Swagger UI)
- **前端 Demo**: http://localhost:8800 (上传+扫描预览)

## 依赖

- **Docker** + **docker compose**
- **Python 3.10+** — `pip install fastapi uvicorn python-multipart pymupdf python-docx`

## API

### `POST /api/convert`
上传 `.docx`，返回转换结果。

```bash
curl -X POST http://localhost:8800/api/convert \
  -F "file=@document.docx"
```

响应:
```json
{
  "id": "a1b2c3d4",
  "fileName": "document.docx",
  "totalPages": 117,
  "pdfUrl": "/api/pdf/a1b2c3d4",
  "pages": ["# page 1 md...", "# page 2 md...", ...]
}
```

### `GET /api/pdf/{id}`
下载转换后的 PDF。

### `GET /api/md/{id}`
获取全部页的 Markdown。

```json
{
  "id": "a1b2c3d4",
  "totalPages": 117,
  "pages": ["...", "..."]
}
```

### `GET /api/md/{id}/{page}`
获取指定页的 Markdown（页码从 1 开始）。

```json
{
  "id": "a1b2c3d4",
  "page": 40,
  "totalPages": 117,
  "markdown": "# page 40 content..."
}
```

### `GET /api/health`
健康检查。

## 工作原理

```
docx 上传
  → ONLYOFFICE ConvertService (Docker) 转 PDF
  → PyMuPDF 提取每页 Markdown (文本块+表格混合)
  → 返回 PDF URL + 逐页 MD
```

## 文件结构

```
docscan/
├── api.py          # FastAPI 服务
├── index.html      # 前端预览 Demo
├── start.sh        # 启动脚本 (自动配置 ONLYOFFICE)
├── stop.sh         # 停止脚本
├── restart.sh      # 重启脚本
├── .gitignore
└── README.md
```
