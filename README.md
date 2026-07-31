# DocScan — Word 文档转 PDF + Markdown API

将 `.docx` 文档精确转换为 **PDF**（保留排版/字体/表格）并提取逐页 **Markdown**（含表格语义）。

## 快速开始

**首次使用**（一键装齐 docker compose / pandoc / Python 依赖 / docker 组权限，再启动）：
```bash
./setup.sh          # 安装 + 启动（默认 8800），首次会拉取 ONLYOFFICE (~3GB)
./setup.sh 8080     # 指定端口
```

**之后日常使用**：
```bash
./start.sh          # 启动服务（无需重新安装）
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

> 除 **Docker Engine** 外，其余依赖均可由 `./setup.sh` 自动安装（docker compose、pandoc、Python 包、docker 组权限）。

- **Docker Engine** — 需预先安装（[官方指南](https://docs.docker.com/engine/install/)）；`docker compose`、中文字体由本项目脚本处理
- **Python 3.10+** — `pip install -r requirements.txt`（setup 自动）
- **pandoc** — md→docx 转换（setup 自动 `apt install pandoc`）

## 配置（可选）

在项目根创建 `docscan.env`（已被 `.gitignore` 忽略），写入需要覆盖默认值的环境变量，`start.sh` 启动时自动加载，改后 `./restart.sh` 生效。例如设置产物保留 24 小时：

```bash
DOCSCAN_RETENTION_HOURS=24
```

| 变量 | 默认 | 作用 |
|---|---|---|
| `DOCSCAN_API_KEY` | 自动生成 | 固定 API Key（不设则存于 `.docscan-api-key`） |
| `DOCSCAN_CORS_ORIGINS` | 空(同源) | CORS 允许域，逗号分隔；公网部署设为具体域 |
| `DOCSCAN_MAX_UPLOAD_MB` | `100` | 单文件上传上限（超限 413） |
| `DOCSCAN_RETENTION_HOURS` | `0`(关) | 启动时清理超过 N 小时的产物 |
| `DOCSCAN_CONVERT_CONCURRENCY` | `4` | ONLYOFFICE 转换并发上限 |
| `DOCSCAN_MAX_QUEUED` | `10` | 排队等待上限（背压）：执行中 + 排队超过此数立即 503 |

查看运行中进程实际生效的配置：`./status.sh`。

## API

> **认证**：除 `GET /api/health`、`/api/docs`、`/openapi.json` 外，所有 `/api/*` 接口都需要 API Key——通过请求头 `X-API-Key: <key>` 或 `Authorization: Bearer <key>` 传入。
> Key 由 `start.sh` 首次启动时自动生成并保存到 `.docscan-api-key`（也可用环境变量 `DOCSCAN_API_KEY` 固定），启动时会打印在终端；前端 Demo 首次访问会弹窗要求输入。下方 curl 示例为简洁省略了该头，实际调用请加上 `-H "X-API-Key: <你的key>"`。

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

### `POST /api/docx/{id}/apply-style`
上传一个**样本 `.docx`**（格式模板），把样本里的**标题样式**（heading 1–4，按大纲级别 `outlineLvl` 跨文档配对）和**正文样式**（`Normal` / `Body Text`）套用到目标文档（`{id}`），就地写回。同时同步样本的**主题字体**（`theme1.xml` 的 majorFont/minorFont），确保 `majorHAnsi` 等主题字体引用解析一致。

套用前自动备份原文件为 `{id}.docx.bak`（仅首次套用时创建，可回滚）；套用失败则恢复原文件。

**配对机制**：标题按 `outlineLvl`（0/1/2/3…）匹配——不依赖两文档的 styleId 或命名是否相同（样本 `styleId="2"` 的 `heading 1` ↔ 目标 `Heading1`）；正文按语义名（`Normal`/`Body Text`/`段落正文`…）匹配。因此样本与目标的样式体系不同也能正确套用。

```bash
curl -X POST http://localhost:8800/api/docx/e73911e954/apply-style \
  -H "X-API-Key: <你的key>" \
  -F "sample=@template.docx"
```

响应（`applied` 列出实际套用了哪些角色；`themeFontsSynced` / `docDefaultsSynced` 表示主题字体与文档默认字体基准是否同步成功）：
```json
{
  "id": "e73911e954",
  "fileName": "document.docx",
  "docxUrl": "/api/docx/e73911e954",
  "applied": [
    {"role": "heading", "targetId": "Heading1", "sampleId": "2", "sampleName": "heading 1"},
    {"role": "heading", "targetId": "Heading2", "sampleId": "3", "sampleName": "heading 2"},
    {"role": "body", "targetId": "Normal", "sampleId": "1", "sampleName": "Normal"},
    {"role": "body", "targetId": "BodyText", "sampleId": "10", "sampleName": "Body Text"}
  ],
  "themeFontsSynced": true,
  "docDefaultsSynced": true,
  "numberingSynced": true
}
```

**字体同步（三层）**：为完整复现样本字体，引擎同步字体链路的三层——① 各标题/正文样式的 rPr（合并保留目标原有的 rFonts 主题绑定，避免回退默认字体）；② `docDefaults/rPrDefault/rPr/rFonts` 文档默认字体基准（含显式中文字体名，如样本指定「黑体」则目标继承「黑体」）；③ `theme1.xml` 的 majorFont/minorFont（主题字体槽位，如 Calibri/宋体）。

**effective 格式**：样本常以 run 级直接格式（在工具栏改字号）覆盖样式定义——引擎取「样式定义 + 代表段落的直接格式」合并为实际渲染外观再套用，避免只搬定义导致字号/加粗与样本所见不符。

**标题编号套用**：若样本标题采用自动编号（`<w:numPr>`，如 `1.`/`1.1.`/`1.1.1.` 多级编号），引擎把 `numbering.xml` 的 `abstractNum`/`num` 定义搬到目标（id 重映射避免冲突），并将 numPr 提升为样式级绑定（H1–H4 自动编号）。判定规则：每个标题层级的**首个段落**决定该级是否编号，且编号必须从顶层（ol=0）连续——父级无编号时子级编号会断裂，不予套用。样本为文本编号（"第一章"/"一、" 写在正文里）则返回 `numberingSynced: false`，不强加自动编号。
套用后用 `GET /api/docx/{id}` 下载即为新外观文档。常用搭配：`md2docx`（生成内容）→ `apply-style`（统一外观）→ `replace` / `crossref`（填占位符、加交叉引用）。

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
├── style_ops.py    # 样式套用引擎（按 outlineLvl 跨文档配对标题/正文样式 + 主题字体同步）
├── index.html      # 前端预览 Demo
├── setup.sh        # 一键安装 + 启动（装 docker compose/pandoc/依赖/权限）
├── start.sh        # 启动脚本 (自动配置 ONLYOFFICE)
├── stop.sh         # 停止脚本
├── restart.sh      # 重启脚本
├── .gitignore
└── README.md
```
