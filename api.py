#!/usr/bin/env python3
"""
DocScan API  —  FastAPI service for docx → PDF + Markdown conversion.

Quick start:
    python3 api.py
    → API at http://localhost:8800/api/
    → Swagger at http://localhost:8800/api/docs
    → Frontend demo at http://localhost:8800/

Endpoints:
    POST /api/convert              upload .docx → returns {id, totalPages, pdfUrl, mdUrl}
    GET  /api/pdf/{id}             download the generated PDF
    GET  /api/md/{id}              all-page Markdown array
    GET  /api/md/{id}/{page}       single-page Markdown
    GET  /api/conversions          list recent conversions
    GET  /api/health               health check

    POST /api/md2docx                    upload .md → returns {id, fileName, docxUrl}
    GET  /api/docx/{id}                  download the current docx
    GET  /api/docx/{id}/placeholders      list 【...】 placeholders with stable ids
    POST /api/docx/{id}/replace           replace placeholders by id
    GET  /api/docx/{id}/tables            list table structure (for crossref target picking)
    GET  /api/docx/{id}/preview           full-text preview (body paragraphs + tables)
    POST /api/docx/{id}/crossref          bookmark a body keyword + insert a page-number
                                           cross-reference field into a target table cell

The frontend ONLYOFFICE viewer is proxied through this server (same-origin
at /oo/…) so the demo at / also works.
"""

import asyncio, json, os, re, secrets, shutil, subprocess, time, uuid, sys
from pathlib import Path
from urllib.parse import urlparse, urlunparse

import httpx
import uvicorn
from docx import Document
from fastapi import FastAPI, File, UploadFile, HTTPException, Request
from fastapi.responses import FileResponse, JSONResponse, HTMLResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

# ——— reuse the conversion engine from server.py ———
ROOT = Path(__file__).parent.resolve()
sys.path.insert(0, str(ROOT))
from server import (
    _convert_docx_to_pdf,
    _extract_pdf_pages,
    _recalculate_fields_docx,
)
import docx_ops
OO_BACKEND = 'http://localhost:8079'   # ONLYOFFICE Docker container

# ——— data dirs ———
DOCS_DIR = ROOT / 'docs'
PDFS_DIR = ROOT / 'pdfs'
MDS_DIR  = ROOT / 'mds'
DOCX_DIR = ROOT / 'docx_store'   # persistent editable docx (md2docx output, placeholder/crossref edits)
for d in (DOCS_DIR, PDFS_DIR, MDS_DIR, DOCX_DIR):
    d.mkdir(parents=True, exist_ok=True)

conversions = {}   # {id: metadata}   in-memory
docx_docs = {}      # {id: metadata}   in-memory, tracks editable docx (see DOCX_DIR)

# 转换是同步阻塞调用（subprocess + httpx.Client），丢进线程池跑，
# 避免卡住 uvicorn 的单个事件循环；并发数上限防止把 ONLYOFFICE 转换 worker 打爆。
CONVERT_CONCURRENCY = int(os.environ.get('DOCSCAN_CONVERT_CONCURRENCY', '4'))
_convert_semaphore = asyncio.Semaphore(CONVERT_CONCURRENCY)

# ——— upload size cap (stream-read so a huge body can't OOM us) ———
MAX_UPLOAD_BYTES = int(os.environ.get('DOCSCAN_MAX_UPLOAD_MB', '100')) * 1024 * 1024

# ——— retention: drop generated files older than N hours on startup (0 = off) ———
RETENTION_HOURS = int(os.environ.get('DOCSCAN_RETENTION_HOURS', '0'))

async def _read_capped(file):
    """Read an UploadFile in 1MB chunks, rejecting bodies over MAX_UPLOAD_BYTES
    (413) instead of slurping the whole thing into memory unbounded."""
    total, chunks = 0, []
    while True:
        chunk = await file.read(1024 * 1024)
        if not chunk:
            break
        total += len(chunk)
        if total > MAX_UPLOAD_BYTES:
            raise HTTPException(413, f'file too large (max {MAX_UPLOAD_BYTES // (1024 * 1024)}MB)')
        chunks.append(chunk)
    return b''.join(chunks)

def _cleanup_old_files():
    """Delete generated files older than RETENTION_HOURS. Off by default — only
    runs when an operator opts in via DOCSCAN_RETENTION_HOURS, so we never
    silently delete a user's data."""
    if RETENTION_HOURS <= 0:
        return
    cutoff = time.time() - RETENTION_HOURS * 3600
    for d in (PDFS_DIR, MDS_DIR, DOCX_DIR, DOCS_DIR):
        for p in d.glob('*'):
            try:
                if p.is_file() and p.stat().st_mtime < cutoff:
                    p.unlink()
            except OSError:
                pass

_cleanup_old_files()

# ——— API key: required for all /api/* except health/docs. If none is ——————
# ——— configured we mint an ephemeral one and print it, so the server never —
# ——— boots unauthenticated by accident. ————————————————————————————————
API_KEY = os.environ.get('DOCSCAN_API_KEY') or secrets.token_urlsafe(24)
# Comma-separated allow-list; '*' (default) stays permissive — but every /api
# call still needs the key, so a wide CORS no longer means "open".
CORS_ORIGINS = [o.strip() for o in os.environ.get('DOCSCAN_CORS_ORIGINS', '*').split(',') if o.strip()]
if not os.environ.get('DOCSCAN_API_KEY'):
    sys.stderr.write(
        '\n==============================================================\n'
        '  No DOCSCAN_API_KEY set — generated ephemeral key:\n'
        f'      {API_KEY}\n'
        '  Set DOCSCAN_API_KEY to keep it stable across restarts.\n'
        '==============================================================\n\n'
    )

# ═══════════════════════════════════════════════════════════════
#  App
# ═══════════════════════════════════════════════════════════════
app = FastAPI(
    title='DocScan API',
    description='Upload .docx → get PDF + per-page Markdown',
    version='1.0',
    docs_url='/api/docs',
)

# Public paths that skip the API key: /api/health (polled by start.sh) and the
# Swagger/OpenAPI schema endpoints (low-sensitivity, handy to leave open).
_PUBLIC_API_PATHS = ('/api/health', '/api/docs', '/openapi')


def _bearer(header: str) -> str:
    parts = header.split(None, 1)
    return parts[1].strip() if len(parts) == 2 and parts[0].lower() == 'bearer' else ''


@app.middleware('http')
async def _require_api_key(request: Request, call_next):
    """Gate every /api/* route behind DOCSCAN_API_KEY (X-API-Key header or an
    Authorization: Bearer token). Frontend pages and the ONLYOFFICE reverse
    proxy are intentionally left open. Registered before the CORS middleware so
    the latter sits outermost and answers OPTIONS preflights without a key."""
    path = request.url.path
    if path.startswith('/api/') and not path.startswith(_PUBLIC_API_PATHS):
        provided = request.headers.get('x-api-key') or _bearer(request.headers.get('authorization', ''))
        if not provided or provided != API_KEY:
            return JSONResponse({'detail': 'invalid or missing API key'}, status_code=401)
    return await call_next(request)

# Added *after* the key middleware so it wraps it (outermost): CORS preflight
# OPTIONS requests get answered here and never reach the key check.
app.add_middleware(CORSMiddleware, allow_origins=CORS_ORIGINS, allow_methods=['*'], allow_headers=['*'])

# ——— helper ———
def _store(docx_name, pdf_path, pages, page_count):
    fid = uuid.uuid4().hex[:10]
    dest = PDFS_DIR / f'{fid}.pdf'
    shutil.move(str(pdf_path), str(dest))
    (MDS_DIR / f'{fid}.json').write_text(json.dumps(pages, ensure_ascii=False), 'utf-8')
    meta = dict(id=fid, fileName=docx_name, totalPages=page_count,
                pdfUrl=f'/api/pdf/{fid}', mdUrl=f'/api/md/{fid}', pages=pages,
                created=time.strftime('%Y-%m-%dT%H:%M:%SZ', time.gmtime()))
    conversions[fid] = meta
    return meta

# ═══════════════════════════════════════════════════════════════
#  API endpoints
# ═══════════════════════════════════════════════════════════════

@app.get('/api/health')
def health(): return dict(status='ok', service='DocScan API', version='1.0.0')

@app.post('/api/convert')
async def convert(file: UploadFile = File(description='.docx file')):
    if not file.filename or not file.filename.lower().endswith('.docx'):
        raise HTTPException(400, 'Only .docx accepted')
    base = uuid.uuid4().hex[:12]
    dx = DOCS_DIR / f'{base}.docx'
    pf = PDFS_DIR / f'{base}.pdf'
    dx.write_bytes(await _read_capped(file))
    loop = asyncio.get_running_loop()
    try:
        async with _convert_semaphore:
            n = await loop.run_in_executor(None, _convert_docx_to_pdf, dx, pf)
            pages = await loop.run_in_executor(None, _extract_pdf_pages, pf)
        meta = _store(file.filename, pf, pages, n or len(pages))
        dx.unlink(missing_ok=True)
        return JSONResponse(meta)
    except Exception as e:
        dx.unlink(missing_ok=True); pf.unlink(missing_ok=True)
        raise HTTPException(500, f'Conversion failed: {e}')

@app.get('/api/pdf/{fid}')
def pdf(fid: str):
    p = PDFS_DIR / f'{fid}.pdf'
    if not p.exists(): raise HTTPException(404, 'not found')
    return FileResponse(str(p), media_type='application/pdf', filename=f'{fid}.pdf')

@app.get('/api/md/{fid}')
def md_all(fid: str):
    j = MDS_DIR / f'{fid}.json'
    if not j.exists(): raise HTTPException(404, 'not found')
    pages = json.loads(j.read_text('utf-8'))
    return dict(id=fid, totalPages=len(pages), pages=pages,
                fileName=conversions.get(fid, {}).get('fileName',''))

@app.get('/api/md/{fid}/{page}')
def md_page(fid: str, page: int):
    j = MDS_DIR / f'{fid}.json'
    if not j.exists(): raise HTTPException(404, 'not found')
    pages = json.loads(j.read_text('utf-8'))
    if page < 1 or page > len(pages): raise HTTPException(404, f'page {page} out of range')
    return dict(id=fid, page=page, totalPages=len(pages), markdown=pages[page-1])

@app.get('/api/conversions')
def list_conv(): return list(conversions.values())

# ═══════════════════════════════════════════════════════════════
#  md → docx, and docx placeholder/crossref editing
# ═══════════════════════════════════════════════════════════════

def _docx_meta(fid, file_name):
    return dict(id=fid, fileName=file_name, docxUrl=f'/api/docx/{fid}',
                created=time.strftime('%Y-%m-%dT%H:%M:%SZ', time.gmtime()))

def _docx_path(fid):
    p = DOCX_DIR / f'{fid}.docx'
    if not p.exists():
        raise HTTPException(404, 'not found')
    return p

def _md_to_docx_sync(md_path, docx_path):
    """pandoc + docx post-processing. Synchronous and blocking, so callers must
    offload it to a worker thread (see md2docx below)."""
    r = subprocess.run(['pandoc', str(md_path), '-o', str(docx_path), '--standalone'],
                       capture_output=True, text=True, timeout=60)
    if r.returncode != 0:
        raise RuntimeError(r.stderr.strip() or 'pandoc exited non-zero')
    doc = Document(str(docx_path))
    docx_ops.convert_hr_to_page_breaks(doc)
    docx_ops.autofit_tables(doc)
    doc.save(str(docx_path))


@app.post('/api/md2docx')
async def md2docx(file: UploadFile = File(description='.md file')):
    if not file.filename or not file.filename.lower().endswith('.md'):
        raise HTTPException(400, 'Only .md accepted')
    fid = uuid.uuid4().hex[:10]
    md_path = DOCS_DIR / f'{fid}.md'
    docx_path = DOCX_DIR / f'{fid}.docx'
    md_path.write_bytes(await _read_capped(file))
    loop = asyncio.get_running_loop()
    try:
        # pandoc + python-docx are blocking — run them in the thread pool so they
        # don't freeze uvicorn's event loop (and /api/health) for up to the
        # pandoc timeout. Shares the convert concurrency cap as back-pressure.
        async with _convert_semaphore:
            await loop.run_in_executor(None, _md_to_docx_sync, md_path, docx_path)
    except Exception as e:
        md_path.unlink(missing_ok=True)
        docx_path.unlink(missing_ok=True)
        raise HTTPException(500, f'md2docx failed: {e}')
    md_path.unlink(missing_ok=True)
    meta = _docx_meta(fid, file.filename)
    docx_docs[fid] = meta
    return JSONResponse(meta)

@app.post('/api/docx/upload')
async def upload_docx(file: UploadFile = File(description='.docx file')):
    """Accept an existing .docx (e.g. from an external generator like
    generate_docx.js) and register it for editing — placeholder listing /
    replacement, cross-reference insertion, preview, and download — exactly
    as if it had been created by /api/md2docx.

    No conversion is performed; the file is stored as-is.
    """
    if not file.filename or not file.filename.lower().endswith('.docx'):
        raise HTTPException(400, 'Only .docx accepted')
    fid = uuid.uuid4().hex[:10]
    docx_path = DOCX_DIR / f'{fid}.docx'
    docx_path.write_bytes(await _read_capped(file))
    meta = _docx_meta(fid, file.filename)
    docx_docs[fid] = meta
    return JSONResponse(meta)

@app.get('/api/docx/{fid}')
def get_docx(fid: str):
    p = _docx_path(fid)
    return FileResponse(str(p), media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                         filename=f'{fid}.docx', headers={'Cache-Control': 'no-store'})

@app.get('/api/docx/{fid}/placeholders')
def get_placeholders(fid: str):
    p = _docx_path(fid)
    doc = Document(str(p))
    placeholders = [ph.to_dict() for ph in docx_ops.list_placeholders(doc)]
    return dict(id=fid, count=len(placeholders), placeholders=placeholders)

class ReplaceRequest(BaseModel):
    replacements: dict[str, str]   # {placeholder_id: new_text}

@app.post('/api/docx/{fid}/replace')
def replace_placeholders(fid: str, body: ReplaceRequest):
    p = _docx_path(fid)
    doc = Document(str(p))
    count = docx_ops.replace_placeholders(doc, body.replacements)
    doc.save(str(p))
    return dict(id=fid, replaced=count)

@app.get('/api/docx/{fid}/tables')
def get_tables(fid: str):
    p = _docx_path(fid)
    doc = Document(str(p))
    return dict(id=fid, tables=docx_ops.list_tables(doc))

@app.get('/api/docx/{fid}/preview')
def get_preview(fid: str):
    """Lightweight full-text preview of the current docx — body paragraphs
    (the only pool eligible as a crossref keyword source) plus all tables.
    Pure local read, no ONLYOFFICE round-trip, so it's fast enough to call
    after every edit to show the effect immediately.
    """
    p = _docx_path(fid)
    doc = Document(str(p))
    return dict(id=fid,
                paragraphs=docx_ops.list_body_paragraphs(doc),
                tables=docx_ops.list_tables(doc))

class CrossrefRequest(BaseModel):
    keyword: str                    # exact text to locate in the document body
    cellPath: str                   # e.g. "table[13].row[1].cell[2]" — from GET .../tables
    paragraphPath: str | None = None  # e.g. "paragraph[5]", from GET .../preview — required
                                       # when `keyword` occurs in more than one body paragraph

@app.post('/api/docx/{fid}/crossref')
async def add_crossref(fid: str, body: CrossrefRequest):
    p = _docx_path(fid)
    doc = Document(str(p))
    try:
        bookmark = docx_ops.add_page_crossref(doc, body.keyword, body.cellPath, body.paragraphPath)
    except ValueError as e:
        raise HTTPException(400, str(e))
    doc.save(str(p))

    # Bake the real page number into the field via ONLYOFFICE, then swap
    # the recalculated file back in as the canonical stored docx.
    loop = asyncio.get_running_loop()
    recalced = DOCX_DIR / f'{fid}-recalc.docx'
    try:
        async with _convert_semaphore:
            await loop.run_in_executor(None, _recalculate_fields_docx, p, recalced)
        shutil.move(str(recalced), str(p))
    except Exception as e:
        recalced.unlink(missing_ok=True)
        raise HTTPException(500, f'page recalculation failed: {e}')

    return dict(id=fid, bookmark=bookmark, cellPath=body.cellPath)

# ═══════════════════════════════════════════════════════════════
#  Frontend demo  ( / → index.html )
# ═══════════════════════════════════════════════════════════════

@app.get('/')
@app.get('/index.html')
def frontend():
    html = (ROOT / 'index.html').read_text('utf-8')
    return HTMLResponse(html, headers={'Cache-Control':'no-cache'})

@app.get('/edit.html')
def edit_frontend():
    html = (ROOT / 'edit.html').read_text('utf-8')
    return HTMLResponse(html, headers={'Cache-Control':'no-cache'})

# ═══════════════════════════════════════════════════════════════
#  ONLYOFFICE reverse proxy   (/oo/*, /coauthoring/*, etc.)
#  Keeps the viewer at the top of index.html same-origin.
# ═══════════════════════════════════════════════════════════════

PROXY_PREFIXES = ('/oo/', '/coauthoring/', '/sdkjs/', '/web-apps/',
                  '/fonts/', '/dictionaries/', '/cache/', '/doc/')

@app.api_route('/{path:path}', methods=['GET','POST','PUT','DELETE','PATCH','OPTIONS','HEAD'])
async def proxy(request: Request, path: str = ''):
    """Catch-all that forwards ONLYOFFICE resources to the Docker container."""
    full = '/' + path
    if not any(full.startswith(p) for p in PROXY_PREFIXES):
        raise HTTPException(404, f'Not found: {full}')

    # Strip /oo prefix — ONLYOFFICE backend serves from root
    backend_path = full
    if backend_path.startswith('/oo/'):
        backend_path = backend_path[3:]  # /oo/xxx → /xxx
    backend_url = OO_BACKEND + backend_path
    headers = dict(request.headers)
    # strip hop-by-hop
    for h in ('host','connection','transfer-encoding','content-length','content-encoding','accept-encoding'):
        headers.pop(h, None)
    # preserve client Host so docservice generates same-origin URLs
    headers['host'] = request.headers.get('host', 'localhost:8800')

    body = await request.body() if request.method in ('POST','PUT','PATCH') else None
    timeout = httpx.Timeout(90, connect=10)
    async with httpx.AsyncClient(timeout=timeout, follow_redirects=False) as client:
        r = await client.request(request.method, backend_url, headers=headers, content=body)

    # rewrite headers (Location must be /-prefixed, not absolute to 8079)
    resp_headers = {}
    for k, v in r.headers.items():
        kl = k.lower()
        if kl in ('transfer-encoding','content-encoding','connection','content-length'):
            continue
        if kl == 'location':
            v = _rewrite_location(v)
        resp_headers[k] = v

    body = r.content  # read full body (httpx streaming has edge cases)
    return Response(
        content=body,
        status_code=r.status_code,
        headers=resp_headers,
        media_type=r.headers.get('content-type'),
    )

def _rewrite_location(loc: str) -> str:
    """Redirect URLs pointing back to the backend into proxy paths."""
    if loc.startswith('/'):
        return loc  # already relative
    try:
        p = urlparse(loc)
        if p.netloc in ('localhost:8079', '127.0.0.1:8079'):
            return p.path + ('?' + p.query if p.query else '')
    except Exception:
        pass
    return loc

# ═══════════════════════════════════════════════════════════════
#  Main
# ═══════════════════════════════════════════════════════════════
if __name__ == '__main__':
    import argparse
    p = argparse.ArgumentParser()
    p.add_argument('--port', type=int, default=8800, help='listen port')
    p.add_argument('--host', default='0.0.0.0')
    args = p.parse_args()
    uvicorn.run('api:app', host=args.host, port=args.port, reload=False, log_level='info')
