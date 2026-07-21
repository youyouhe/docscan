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
- **Python 3.10+** — `pip install -r requirements.txt`
- **pandoc** — md→docx 转换（`apt install pandoc` 或对应系统包管理器）

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

### `POST /api/md2docx`
上传 `.md`，用 pandoc 转换为 `.docx`（保留标题层级/表格/列表结构）。

```bash
curl -X POST http://localhost:8800/api/md2docx \
  -F "file=@document.md"
```

响应:
```json
{
  "id": "e73911e954",
  "fileName": "document.md",
  "docxUrl": "/api/docx/e73911e954"
}
```

### `POST /api/docx/upload`
上传已有的 `.docx`（如 `generate_docx.js` 输出），注册到 DocScan 供后续编辑——占位符提取/替换、表格列表、交叉引用插入——与 `md2docx` 产出的 docx 一样可编辑。不做任何转换，仅存储。

```bash
curl -X POST http://localhost:8800/api/docx/upload \
  -F "file=@output.docx"
```

响应:
```json
{
  "id": "a1b2c3d4e5",
  "fileName": "output.docx",
  "docxUrl": "/api/docx/a1b2c3d4e5"
}
```

### `GET /api/docx/{id}`
下载当前的 docx（含后续 replace/crossref 的修改）。

### `GET /api/docx/{id}/placeholders`
提取文档正文+表格中所有 `【...】` 占位符，按文档顺序编号（`ph-0`, `ph-1`, ...）。
同一占位符文本（如多个 `【待填写】`）在文档中出现多次时，每次出现都有独立 id。

```json
{
  "id": "e73911e954",
  "count": 62,
  "placeholders": [
    {"id": "ph-0", "text": "【待填：X份，金额均≥40万元】", "location": "table", "path": "table[23].row[1].cell[4]"},
    {"id": "ph-1", "text": "【待填写】", "location": "table", "path": "table[23].row[2].cell[4]"}
  ]
}
```

### `POST /api/docx/{id}/replace`
按占位符 id（不是文本）批量替换，避免同名占位符互相干扰。

```bash
curl -X POST http://localhost:8800/api/docx/e73911e954/replace \
  -H "Content-Type: application/json" \
  -d '{"replacements": {"ph-1": "1250.5", "ph-2": "3800.2"}}'
```

### `GET /api/docx/{id}/tables`
列出文档中所有表格的坐标和单元格文本，用于挑选页码交叉引用的目标单元格。

```json
{
  "id": "e73911e954",
  "tables": [
    {"path": "table[13]", "rows": [["序号", "符合性审查项目", "文件名称/页码"], ["1", "报价文件签字...", "..."]]}
  ]
}
```

### `POST /api/docx/{id}/crossref`
在正文中精确匹配 `keyword` 文本并打书签，在 `cellPath` 指定的表格单元格插入页码交叉引用字段，
并通过 ONLYOFFICE 重新排版计算出真实页码后写回。

```bash
curl -X POST http://localhost:8800/api/docx/e73911e954/crossref \
  -H "Content-Type: application/json" \
  -d '{"keyword": "某单位被装仓储无纸化办公建设项目", "cellPath": "table[13].row[1].cell[2]"}'
```

响应:
```json
{
  "id": "e73911e954",
  "bookmark": "bm_0132ee8176bb",
  "cellPath": "table[13].row[1].cell[2]"
}
```

## 工作原理

```
docx 上传
  → ONLYOFFICE ConvertService (Docker，经网桥网关直接回源取文件) 转 PDF
  → PyMuPDF 提取每页文本，并用 find_tables() 检测表格区域
    → 表格区域按行列渲染为 GFM Markdown 表格（| 分隔），非表格文字按原阅读顺序保留为纯文本
  → 返回 PDF URL + 逐页 MD
```

**表格识别的边界**：`find_tables()` 基于版式（网格线/对齐）检测表格，对绝大多数规整表格（评分表、报价表、索引表等）能准确还原行列结构；但极少数复杂排版（如单元格内嵌套小表格、无边框纯空格对齐的伪表格）可能识别不到或识别有误。下游读取 Markdown 时如遇到关键数据（金额、分值）所在段落不像表格，应结合原 PDF 或 Word 交叉核对，不要仅凭本服务输出的 Markdown 下结论。

md → docx 编辑流程：
```
md 上传 → pandoc 转 docx
  → python-docx 提取/替换 【占位符】(按位置定位，不按文本)
  → python-docx 在正文打书签 + 表格单元格插入 PAGEREF 字段
  → ONLYOFFICE DocBuilder 重新排版计算真实页码(ForceRecalculate+UpdateAllFields)
  → 写回 docx
```

## 文件结构

```
docscan/
├── api.py          # FastAPI 服务
├── server.py       # docx→pdf 转换引擎（ONLYOFFICE）+ PDF→表格感知 Markdown（PyMuPDF find_tables）+ 字段重算引擎
├── docx_ops.py     # 占位符提取/替换、书签/页码交叉引用（python-docx 直操 XML）
├── index.html      # 前端预览 Demo
├── start.sh        # 启动脚本 (自动配置 ONLYOFFICE)
├── stop.sh         # 停止脚本
├── restart.sh      # 重启脚本
├── .gitignore
└── README.md
```
