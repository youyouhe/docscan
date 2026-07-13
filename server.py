#!/usr/bin/env python3
"""
DocScan conversion engine  —  the layer under api.py.

api.py is only an HTTP shell; the real work lives here. Three functions are
imported by api.py:

    _ensure_container_file_server()   make sure ONLYOFFICE can fetch our files
    _convert_docx_to_pdf(docx, pdf)   docx -> PDF via ONLYOFFICE; returns page count
    _extract_pdf_pages(pdf)           per-page text via PyMuPDF; returns list[str]

Conversion path (the ONLYOFFICE container mounts no host dir, so we ferry files
through it):

    host docx
      -- docker cp -->  container:/tmp/<base>.docx
      <-- served by -->  in-container `python3 -m http.server 9999` (loopback)
      <-- ONLYOFFICE --> POST /converter  (url = http://localhost:9999/<base>.docx)
      --> ONLYOFFICE writes the PDF to its cache and replies with a host-facing
          fileUrl on port 8079 (http://localhost:8079/cache/files/...)
    host pdf  <-- GET that fileUrl

All endpoints/behaviours below were pinned by live probing against the running
onlyoffice/documentserver:latest container (JWT disabled).
"""

import subprocess
import time
from pathlib import Path

import httpx

try:
    import fitz  # PyMuPDF
except ImportError:  # very new pymupdf may drop the fitz alias
    import pymupdf as fitz

OO_BACKEND = 'http://localhost:8079'   # ONLYOFFICE container, host-facing port
CONTAINER = 'onlyoffice'                # docker container name
FS_PORT = 9999                          # in-container file server (loopback only)
CONVERT_PATH = '/converter'             # ONLYOFFICE 8.x ConvertService endpoint
CONVERT_TIMEOUT = 180.0                 # seconds, large docs can take a while


def _run(cmd):
    """Run a command list, capturing output. Returns CompletedProcess (no raise)."""
    return subprocess.run(cmd, capture_output=True, text=True)


# ════════════════════════════════════════════════════════════════════
#  In-container file server
# ════════════════════════════════════════════════════════════════════
def _ensure_container_file_server():
    """Ensure ONLYOFFICE can fetch files we push into the container.

    Idempotent: if :9999 already answers inside the container, return immediately.
    Otherwise launch `python3 -m http.server 9999` detached in /tmp and wait for
    it to come up. start.sh starts this too; we re-assert so api.py works even
    when launched standalone.
    """
    url = f'http://localhost:{FS_PORT}/'
    if _container_http_ok(url):
        return True

    _run(['docker', 'exec', '-d', CONTAINER, 'sh', '-c',
          f'cd /tmp && nohup python3 -m http.server {FS_PORT} >/tmp/fs.log 2>&1 &'])

    for _ in range(50):  # wait up to ~5s
        if _container_http_ok(url):
            return True
        time.sleep(0.1)
    raise RuntimeError(f'container file server on :{FS_PORT} did not come up')


def _container_http_ok(url):
    """True if curl inside the container gets a 2xx/3xx for url."""
    r = subprocess.run(
        ['docker', 'exec', CONTAINER, 'curl', '-s', '-o', '/dev/null',
         '-w', '%{http_code}', '--max-time', '3', url],
        capture_output=True, text=True)
    code = (r.stdout or '').strip()
    return bool(code) and code[0] in '23'


# ════════════════════════════════════════════════════════════════════
#  docx -> PDF via ONLYOFFICE ConvertService
# ════════════════════════════════════════════════════════════════════
def _convert_docx_to_pdf(docx_path, pdf_path, *, key=None):
    """Convert a host .docx to a host PDF via ONLYOFFICE.

    Returns the PDF page count (int; 0 if it can't be read). Raises on any
    conversion or download failure. Cleans up the pushed source regardless.
    """
    docx_path, pdf_path = Path(docx_path), Path(pdf_path)
    base = docx_path.stem
    key = key or f'{base}-{int(time.time())}'   # unique key => never a stale cache hit

    _ensure_container_file_server()
    _cp_into_container(docx_path, f'/tmp/{base}.docx')
    src_url = f'http://localhost:{FS_PORT}/{base}.docx'

    payload = {
        'async': False,
        'filetype': 'docx',
        'key': key,
        'outputtype': 'pdf',
        'title': f'{base}.docx',
        'url': src_url,
    }
    try:
        file_url = _request_convert(payload)   # host-facing, :8079
        _download(file_url, pdf_path)
    finally:
        _run(['docker', 'exec', CONTAINER, 'rm', '-f', f'/tmp/{base}.docx'])

    return _pdf_page_count(pdf_path)


def _request_convert(payload):
    """POST to ONLYOFFICE ConvertService; return the host-reachable result fileUrl."""
    key = payload['key']
    with httpx.Client(timeout=CONVERT_TIMEOUT) as c:
        r = c.post(OO_BACKEND + CONVERT_PATH, params={'shardkey': key}, json=payload)
        r.raise_for_status()
        data = r.json()
    if not data.get('endConvert'):
        raise RuntimeError(f'conversion did not finish: {data}')
    file_url = data.get('fileUrl')
    if not file_url:
        raise RuntimeError(f'no fileUrl in conversion response: {data}')
    return file_url


def _download(file_url, dest):
    """Stream-download file_url (host-facing, port 8079) to dest path."""
    with httpx.Client(timeout=CONVERT_TIMEOUT) as c, c.stream('GET', file_url) as r:
        r.raise_for_status()
        with open(dest, 'wb') as f:
            for chunk in r.iter_bytes():
                f.write(chunk)


def _cp_into_container(src, container_path):
    r = _run(['docker', 'cp', str(src), f'{CONTAINER}:{container_path}'])
    if r.returncode != 0:
        raise RuntimeError(f'docker cp failed: {r.stderr.strip()}')


def _pdf_page_count(pdf_path):
    try:
        with fitz.open(str(pdf_path)) as d:
            return d.page_count
    except Exception:
        return 0


# ════════════════════════════════════════════════════════════════════
#  PDF -> per-page text (Markdown-ish, table-aware)
# ════════════════════════════════════════════════════════════════════
def _extract_pdf_pages(pdf_path):
    """Extract per-page Markdown from a PDF; return list[str] (one string per
    page, ordered, 1-indexed by position). Pages that are image-only yield ''.

    Tables are detected via PyMuPDF's find_tables() and rendered as GFM
    pipe-tables so row/column structure survives (scoring tables, quotation
    tables, etc.). Non-table text is extracted as plain text blocks. Both are
    interleaved in top-to-bottom reading order using each block's bbox.
    """
    pages = []
    with fitz.open(str(pdf_path)) as doc:
        for page in doc:
            pages.append(_tidy(_page_to_markdown(page)))
    return pages


def _drop_nested_tables(tables, tol=2.0):
    """find_tables() sometimes misdetects a complex cell inside a big table as
    a second, independent table fully contained within the first one's bbox.
    Drop any table whose bbox is contained within another's to avoid emitting
    the same region twice.
    """
    def _contains(outer, inner):
        ox0, oy0, ox1, oy1 = outer
        ix0, iy0, ix1, iy1 = inner
        return (ix0 >= ox0 - tol and iy0 >= oy0 - tol and
                ix1 <= ox1 + tol and iy1 <= oy1 + tol)

    keep = []
    for i, t in enumerate(tables):
        nested_in_other = any(
            i != j and _contains(other.bbox, t.bbox) and other.bbox != t.bbox
            for j, other in enumerate(tables)
        )
        if not nested_in_other:
            keep.append(t)
    return keep


def _page_to_markdown(page):
    """Render one page as text, with detected tables rendered as Markdown
    pipe-tables, interleaved with surrounding text by vertical position.
    """
    try:
        tables = page.find_tables().tables
    except Exception:
        tables = []
    tables = _drop_nested_tables(tables)

    if not tables:
        return page.get_text('text') or ''

    # Text blocks (x0, y0, x1, y1, text, block_no, block_type)
    text_blocks = [b for b in page.get_text('blocks') if b[6] == 0]

    def _overlaps_table(block, table_bbox, tol=2.0):
        bx0, by0, bx1, by1 = block[:4]
        tx0, ty0, tx1, ty1 = table_bbox
        return not (bx1 < tx0 - tol or bx0 > tx1 + tol or by1 < ty0 - tol or by0 > ty1 + tol)

    items = []  # (y0, kind, payload)
    for t in tables:
        items.append((t.bbox[1], 'table', t))
    for b in text_blocks:
        if any(_overlaps_table(b, t.bbox) for t in tables):
            continue  # skip text that's actually inside a table region
        txt = (b[4] or '').strip()
        if txt:
            items.append((b[1], 'text', txt))

    items.sort(key=lambda it: it[0])

    parts = []
    for _, kind, payload in items:
        if kind == 'text':
            parts.append(payload)
        else:
            md = _table_to_markdown(payload)
            if md:
                parts.append(md)
    return '\n\n'.join(parts)


def _table_to_markdown(table):
    """Render a fitz Table as a GFM pipe-table. Returns '' if extraction fails
    or the table has no usable rows.
    """
    try:
        rows = table.extract()
    except Exception:
        return ''
    if not rows:
        return ''

    def _cell(v):
        return (v or '').replace('\n', ' ').replace('|', '\\|').strip()

    header = [_cell(c) for c in rows[0]]
    ncol = len(header) or 1
    lines = ['| ' + ' | '.join(header) + ' |',
              '|' + '|'.join(['---'] * ncol) + '|']
    for row in rows[1:]:
        cells = [_cell(c) for c in row]
        cells += [''] * (ncol - len(cells))  # pad short rows
        lines.append('| ' + ' | '.join(cells[:ncol]) + ' |')
    return '\n'.join(lines)


def _tidy(text):
    """Collapse runs of blank lines, trim edges."""
    out, blanks = [], 0
    for ln in text.splitlines():
        if ln.strip() == '':
            blanks += 1
            if blanks <= 1:
                out.append('')
        else:
            blanks = 0
            out.append(ln.rstrip())
    return '\n'.join(out).strip()


# ════════════════════════════════════════════════════════════════════
#  Manual smoke test:  python3 server.py <file.docx> [out.pdf]
# ════════════════════════════════════════════════════════════════════
if __name__ == '__main__':
    import sys, tempfile
    if len(sys.argv) < 2:
        print('usage: server.py <file.docx> [out.pdf]')
        sys.exit(1)
    src = Path(sys.argv[1])
    out = Path(sys.argv[2]) if len(sys.argv) > 2 \
        else Path(tempfile.gettempdir()) / f'{src.stem}.pdf'
    n = _convert_docx_to_pdf(src, out)
    pages = _extract_pdf_pages(out)
    print(f'PDF  : {out}')
    print(f'pages: {n}  (extracted {len(pages)})')
    for i, md in enumerate(pages[:2], 1):
        print(f'\n--- page {i} (first 240 chars) ---\n{md[:240]}')
