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

The frontend ONLYOFFICE viewer is proxied through this server (same-origin
at /oo/…) so the demo at / also works.
"""

import json, os, re, shutil, time, uuid, sys
from pathlib import Path
from urllib.parse import urlparse, urlunparse

import httpx
import uvicorn
from fastapi import FastAPI, File, UploadFile, HTTPException, Request
from fastapi.responses import FileResponse, JSONResponse, HTMLResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware

# ——— reuse the conversion engine from server.py ———
ROOT = Path(__file__).parent.resolve()
sys.path.insert(0, str(ROOT))
from server import (
    _convert_docx_to_pdf,
    _extract_pdf_pages,
    _ensure_container_file_server,
)
OO_BACKEND = 'http://localhost:8079'   # ONLYOFFICE Docker container

# ——— data dirs ———
DOCS_DIR = ROOT / 'docs'
PDFS_DIR = ROOT / 'pdfs'
MDS_DIR  = ROOT / 'mds'
for d in (DOCS_DIR, PDFS_DIR, MDS_DIR):
    d.mkdir(parents=True, exist_ok=True)

conversions = {}   # {id: metadata}   in-memory

# ═══════════════════════════════════════════════════════════════
#  App
# ═══════════════════════════════════════════════════════════════
app = FastAPI(
    title='DocScan API',
    description='Upload .docx → get PDF + per-page Markdown',
    version='1.0',
    docs_url='/api/docs',
)
app.add_middleware(CORSMiddleware, allow_origins=['*'], allow_methods=['*'], allow_headers=['*'])

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
    dx.write_bytes(await file.read())
    try:
        n = _convert_docx_to_pdf(dx, pf)
        pages = _extract_pdf_pages(pf)
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
#  Frontend demo  ( / → index.html )
# ═══════════════════════════════════════════════════════════════

@app.get('/')
@app.get('/index.html')
def frontend():
    html = (ROOT / 'index.html').read_text('utf-8')
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
