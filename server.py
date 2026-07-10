#!/usr/bin/env python3
"""
DocScan conversion engine  —  the layer under api.py.

api.py is only an HTTP shell; the real work lives here. Three functions are
imported by api.py:

    _ensure_container_file_server()   no-op kept for backwards compatibility
    _convert_docx_to_pdf(docx, pdf)   docx -> PDF via ONLYOFFICE; returns page count
    _extract_pdf_pages(pdf)           per-page text via PyMuPDF; returns list[str]

Conversion path (ONLYOFFICE reaches the host over the Docker bridge network,
so we serve the docx straight off the host — no docker cp / container-side
file server needed):

    host docx bytes
      -- served by -->  one-shot ThreadingHTTPServer bound to the docker
                         bridge gateway IP, random port, single-file only
      <-- ONLYOFFICE --> POST /converter  (url = http://<gateway>:<port>/<base>.docx)
      --> ONLYOFFICE writes the PDF to its cache and replies with a host-facing
          fileUrl on port 8079 (http://localhost:8079/cache/files/...)
    host pdf  <-- GET that fileUrl

Each conversion gets its own ephemeral server/port, so concurrent conversions
never share state. The gateway IP is auto-detected from the running container
(cached after first lookup) rather than hardcoded, since it depends on which
Docker network the container ends up on.

All endpoints/behaviours below were pinned by live probing against the running
onlyoffice/documentserver:latest container (JWT disabled).
"""

import http.server
import subprocess
import threading
import time
from contextlib import contextmanager
from pathlib import Path

import httpx

try:
    import fitz  # PyMuPDF
except ImportError:  # very new pymupdf may drop the fitz alias
    import pymupdf as fitz

OO_BACKEND = 'http://localhost:8079'   # ONLYOFFICE container, host-facing port
CONTAINER = 'onlyoffice'                # docker container name
CONVERT_PATH = '/converter'             # ONLYOFFICE 8.x ConvertService endpoint
DOCBUILDER_PATH = '/docbuilder'         # ONLYOFFICE DocBuilder endpoint — body is the
                                        # raw .docbuilder script text, NOT JSON-wrapped
CONVERT_TIMEOUT = 180.0                 # seconds, large docs can take a while

_gateway_ip_cache = None


def _run(cmd):
    """Run a command list, capturing output. Returns CompletedProcess (no raise)."""
    return subprocess.run(cmd, capture_output=True, text=True)


def _ensure_container_file_server():
    """Kept as a no-op so older callers/imports don't break; nothing to start."""
    return True


# ════════════════════════════════════════════════════════════════════
#  Docker bridge gateway IP  (host address reachable *from* the container)
# ════════════════════════════════════════════════════════════════════
def _container_gateway_ip():
    """Host IP the ONLYOFFICE container can reach us on, cached after first lookup.

    Detected from the container's own network config rather than hardcoded,
    since it depends on which Docker network / compose project started it.
    """
    global _gateway_ip_cache
    if _gateway_ip_cache:
        return _gateway_ip_cache

    r = _run(['docker', 'inspect', CONTAINER, '--format',
              '{{range $k, $v := .NetworkSettings.Networks}}{{$v.Gateway}}{{"\\n"}}{{end}}'])
    if r.returncode != 0:
        raise RuntimeError(f'docker inspect {CONTAINER} failed: {r.stderr.strip()}')
    ips = [ln.strip() for ln in (r.stdout or '').splitlines() if ln.strip()]
    if not ips:
        raise RuntimeError(f'could not determine gateway IP for container {CONTAINER}')

    _gateway_ip_cache = ips[0]
    return _gateway_ip_cache


# ════════════════════════════════════════════════════════════════════
#  One-shot single-file HTTP server (host -> container, per conversion)
# ════════════════════════════════════════════════════════════════════
class _SingleFileHandler(http.server.BaseHTTPRequestHandler):
    def __init__(self, *args, data=b'', file_name='', **kwargs):
        self._data = data
        self._file_name = file_name
        super().__init__(*args, **kwargs)

    def do_GET(self):
        if self.path.lstrip('/') != self._file_name:
            self.send_error(404)
            return
        self.send_response(200)
        self.send_header('Content-Type', 'application/octet-stream')
        self.send_header('Content-Length', str(len(self._data)))
        self.end_headers()
        self.wfile.write(self._data)

    def log_message(self, fmt, *args):  # silence per-request stderr noise
        pass


@contextmanager
def _serve_file_once(file_name, data):
    """Serve `data` at /<file_name> on a random port bound to the docker
    gateway IP, for the lifetime of the `with` block. One ephemeral server
    per call, so concurrent conversions never share a port or file.
    """
    from functools import partial
    handler = partial(_SingleFileHandler, data=data, file_name=file_name)
    bind_ip = _container_gateway_ip()
    httpd = http.server.ThreadingHTTPServer((bind_ip, 0), handler)
    thread = threading.Thread(target=httpd.serve_forever, daemon=True)
    thread.start()
    try:
        yield bind_ip, httpd.server_address[1]
    finally:
        httpd.shutdown()
        httpd.server_close()


# ════════════════════════════════════════════════════════════════════
#  docx -> PDF via ONLYOFFICE ConvertService
# ════════════════════════════════════════════════════════════════════
def _convert_docx_to_pdf(docx_path, pdf_path, *, key=None):
    """Convert a host .docx to a host PDF via ONLYOFFICE.

    Returns the PDF page count (int; 0 if it can't be read). Raises on any
    conversion or download failure.
    """
    docx_path, pdf_path = Path(docx_path), Path(pdf_path)
    base = docx_path.stem
    key = key or f'{base}-{int(time.time())}'   # unique key => never a stale cache hit
    file_name = f'{base}.docx'
    data = docx_path.read_bytes()

    with _serve_file_once(file_name, data) as (bind_ip, port):
        src_url = f'http://{bind_ip}:{port}/{file_name}'
        payload = {
            'async': False,
            'filetype': 'docx',
            'key': key,
            'outputtype': 'pdf',
            'title': file_name,
            'url': src_url,
        }
        file_url = _request_convert(payload)   # host-facing, :8079
        _download(file_url, pdf_path)

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


def _pdf_page_count(pdf_path):
    try:
        with fitz.open(str(pdf_path)) as d:
            return d.page_count
    except Exception:
        return 0


# ════════════════════════════════════════════════════════════════════
#  Recalculate fields (e.g. PAGEREF) via ONLYOFFICE DocBuilder
# ════════════════════════════════════════════════════════════════════
def _recalculate_fields_docx(docx_path, out_path):
    """Force ONLYOFFICE to lay out the document and bake real values into
    field codes (e.g. PAGEREF page numbers), then save. Needed because
    fields inserted by python-docx (or by us) carry a cached display value
    of 0/blank until something actually paginates the document.

    The DocBuilder script here is a fixed template — no user data is ever
    interpolated into it, only the source/output file names we control —
    so there's no script-injection surface even though the endpoint takes
    raw JS text as its POST body.
    """
    docx_path, out_path = Path(docx_path), Path(out_path)
    base = docx_path.stem
    file_name = f'{base}.docx'
    out_name = f'{base}-recalc.docx'
    data = docx_path.read_bytes()

    with _serve_file_once(file_name, data) as (bind_ip, port):
        src_url = f'http://{bind_ip}:{port}/{file_name}'
        script = (
            f'builder.OpenFile("{src_url}", "docx");\n'
            'var oDocument = Api.GetDocument();\n'
            'oDocument.ForceRecalculate();\n'
            'oDocument.UpdateAllFields();\n'
            'oDocument.ForceRecalculate();\n'
            f'builder.SaveFile("docx", "{out_name}");\n'
            'builder.CloseFile();\n'
        )
        file_url = _request_docbuilder(script)
        _download(file_url, out_path)


def _request_docbuilder(script):
    """POST a raw DocBuilder script to ONLYOFFICE; return the first output fileUrl.

    Unlike ConvertService, this endpoint expects the script text itself as
    the POST body — wrapping it in {"script": ...} JSON makes ONLYOFFICE try
    to parse the wrapper as JS and fail with a syntax error.
    """
    with httpx.Client(timeout=CONVERT_TIMEOUT) as c:
        r = c.post(OO_BACKEND + DOCBUILDER_PATH,
                   headers={'Content-Type': 'application/json'},
                   content=script.encode('utf-8'))
        r.raise_for_status()
        data = r.json()
    if data.get('error'):
        raise RuntimeError(f'docbuilder error: {data}')
    urls = data.get('urls') or {}
    if not urls:
        raise RuntimeError(f'no output urls from docbuilder: {data}')
    return next(iter(urls.values()))


# ════════════════════════════════════════════════════════════════════
#  PDF -> per-page text (Markdown-ish)
# ════════════════════════════════════════════════════════════════════
def _extract_pdf_pages(pdf_path):
    """Extract text per page from a PDF; return list[str] (one string per page,
    ordered, 1-indexed by position). Pages that are image-only yield ''.
    """
    pages = []
    with fitz.open(str(pdf_path)) as doc:
        for page in doc:
            pages.append(_tidy(page.get_text('text') or ''))
    return pages


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
