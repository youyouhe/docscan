#!/usr/bin/env python3
"""
DocScan docx editing engine — placeholder extraction/replace and page-number
cross-references, operating directly on docx XML via python-docx.

Three entry points, all pure functions taking/returning python-docx Document
objects (callers own load/save):

    list_placeholders(doc)                     -> list[Placeholder]
    replace_placeholders(doc, {id: value})      -> int (count replaced)
    list_tables(doc)                            -> list[TableInfo]
    add_page_crossref(doc, keyword, cell_path, paragraph_path=None)  -> bookmark name
    list_body_paragraphs(doc)                   -> list[str]  (for previewing)
    autofit_tables(doc, page_width=None)        -> int (count of tables resized)
    convert_hr_to_page_breaks(doc)              -> int (count of `---` converted to page breaks)

Design note: replacement is done by *placeholder id*, not by placeholder
text, because the same literal text (e.g. "【待填写】") recurs many times
across a real document — a text->value dict can't disambiguate which
occurrence gets which value. Ids are assigned by document order.
"""

import re
import unicodedata
import uuid
from docx.oxml.ns import qn
from docx.shared import Twips
from docx.table import Table
from docx.text.paragraph import Paragraph

PLACEHOLDER_RE = re.compile(r'【[^【】]*】')


class Placeholder:
    def __init__(self, id, text, location, path):
        self.id = id
        self.text = text
        self.location = location  # 'body' | 'table'
        self.path = path          # e.g. "table[2].row[1].cell[3]" or "paragraph[5]"

    def to_dict(self):
        return dict(id=self.id, text=self.text, location=self.location, path=self.path)


# ════════════════════════════════════════════════════════════════════
#  Run-level helpers — a placeholder's 【...】 text may be split across
#  multiple <w:r> runs (whatever the source editor happened to produce),
#  so matching/replacing has to work on the paragraph's concatenated text
#  and then map back to the runs it spans.
# ════════════════════════════════════════════════════════════════════
def _iter_paragraphs_with_path(doc):
    """Yield (paragraph, path_str) for every paragraph in the body and in
    every table cell, in document order. Nested tables (rare, but legal)
    are descended into as well.
    """
    for i, item in enumerate(doc.iter_inner_content()):
        if isinstance(item, Paragraph):
            yield item, f'paragraph[{i}]'
        elif isinstance(item, Table):
            yield from _iter_table_paragraphs(item, f'table[{i}]')


def _iter_table_paragraphs(table, table_path):
    for r, row in enumerate(table.rows):
        for c, cell in enumerate(row.cells):
            cell_path = f'{table_path}.row[{r}].cell[{c}]'
            for item in cell.iter_inner_content():
                if isinstance(item, Paragraph):
                    yield item, cell_path
                elif isinstance(item, Table):
                    yield from _iter_table_paragraphs(item, f'{cell_path}.table')


def _find_matches(paragraph):
    """Return [(match_text, start_char, end_char)] for every 【...】 span
    in this paragraph's full text (concatenation of all its runs).
    """
    text = paragraph.text
    return [(m.group(0), m.start(), m.end()) for m in PLACEHOLDER_RE.finditer(text)]


def _run_spans(paragraph):
    """Return [(run, start_char, end_char)] mapping each run to the char
    range it occupies in paragraph.text.
    """
    spans = []
    pos = 0
    for run in paragraph.runs:
        length = len(run.text)
        spans.append((run, pos, pos + length))
        pos += length
    return spans


def _replace_span(paragraph, start, end, new_text):
    """Replace paragraph.text[start:end] with new_text, updating whichever
    runs the span touches. The first touched run keeps its formatting and
    receives the new text; other touched runs are emptied (not removed —
    removing radically changes run/rPr bookkeeping for little benefit here).
    """
    spans = _run_spans(paragraph)
    first = True
    for run, r_start, r_end in spans:
        overlap_start = max(start, r_start)
        overlap_end = min(end, r_end)
        if overlap_start >= overlap_end:
            continue  # this run isn't touched by [start, end)
        local_start = overlap_start - r_start
        local_end = overlap_end - r_start
        before = run.text[:local_start]
        after = run.text[local_end:]
        if first:
            run.text = before + new_text + after
            first = False
        else:
            run.text = before + after


# ════════════════════════════════════════════════════════════════════
#  Placeholder listing / replacement
# ════════════════════════════════════════════════════════════════════
def list_placeholders(doc):
    """Find every 【...】 span in the document (body + tables, nested tables
    included), in document order. Each gets a stable-for-this-call id
    ("ph-0", "ph-1", ...) reflecting that order.
    """
    out = []
    idx = 0
    for paragraph, path in _iter_paragraphs_with_path(doc):
        for text, start, end in _find_matches(paragraph):
            location = 'table' if '.cell[' in path else 'body'
            out.append(Placeholder(f'ph-{idx}', text, location, path))
            idx += 1
    return out


def replace_placeholders(doc, replacements):
    """Replace placeholders by id. `replacements` is {placeholder_id: new_text}.
    Returns the number of placeholders actually replaced. Ids not present in
    the document are silently ignored (caller already has the fresh id list
    from list_placeholders, so a mismatch means the document changed underneath).
    """
    if not replacements:
        return 0
    count = 0
    idx = 0
    for paragraph, _path in _iter_paragraphs_with_path(doc):
        matches = _find_matches(paragraph)
        if not matches:
            continue
        # Replace back-to-front so earlier spans' char offsets in this
        # paragraph stay valid while later ones are still being replaced.
        for offset in reversed(range(len(matches))):
            text, start, end = matches[offset]
            ph_id = f'ph-{idx + offset}'
            if ph_id in replacements:
                _replace_span(paragraph, start, end, replacements[ph_id])
                count += 1
        idx += len(matches)
    return count


# ════════════════════════════════════════════════════════════════════
#  Table structure listing (for choosing a cross-ref target cell)
# ════════════════════════════════════════════════════════════════════
def list_tables(doc):
    """Return [{path, rows: [[cell_text, ...], ...]}] for every top-level
    table in document order. Nested tables aren't included here — this is
    meant for picking a target cell for a page-number cross-reference,
    which in practice always lands in a top-level index table.
    """
    out = []
    for i, item in enumerate(doc.iter_inner_content()):
        if not isinstance(item, Table):
            continue
        rows = [[cell.text for cell in row.cells] for row in item.rows]
        out.append(dict(path=f'table[{i}]', rows=rows))
    return out


def list_body_paragraphs(doc):
    """Return [{path, text}] for every top-level body paragraph (not inside
    tables), in document order, skipping blank ones. This is the pool that
    add_page_crossref's `keyword` is matched against — useful for a UI to
    show the operator what text is actually eligible as a crossref source,
    since a keyword that only appears inside a table will be rejected.
    """
    out = []
    for i, item in enumerate(doc.iter_inner_content()):
        if isinstance(item, Paragraph) and item.text.strip():
            out.append(dict(path=f'paragraph[{i}]', text=item.text))
    return out


# ════════════════════════════════════════════════════════════════════
#  Horizontal-rule -> page-break — pandoc renders markdown's `---` as a
#  VML rect ("draw a thin line"), never as a real page break. We detect
#  that marker and swap the paragraph's content for <w:br w:type="page"/>,
#  since a manual line of `---` in these report templates is almost always
#  meant to separate sections (cover page vs. body, one annex vs. the next)
#  rather than to visually rule off two paragraphs.
# ════════════════════════════════════════════════════════════════════
def convert_hr_to_page_breaks(doc):
    """Replace every pandoc horizontal-rule paragraph with a page break.
    Returns the number of paragraphs converted.
    """
    from docx.oxml import OxmlElement
    count = 0
    for item in list(doc.iter_inner_content()):
        if not isinstance(item, Paragraph):
            continue
        p = item._p
        if not p.xpath(".//*[local-name()='rect'][@*[local-name()='hr']='t']"):
            continue
        p.remove_all('w:r', 'w:hyperlink', 'w:bookmarkStart', 'w:bookmarkEnd')
        run = OxmlElement('w:r')
        br = OxmlElement('w:br')
        br.set(qn('w:type'), 'page')
        run.append(br)
        p.append(run)
        count += 1
    return count


# ════════════════════════════════════════════════════════════════════
#  Table column autofit — pandoc emits tables with tblW type="pct" w="0.0"
#  and often an empty tblGrid (no column widths at all), which ONLYOFFICE's
#  PDF renderer turns into cramped, evenly-split columns regardless of
#  content. We replicate Word's "AutoFit to Contents" by measuring each
#  column's text and writing explicit fixed-layout widths.
# ════════════════════════════════════════════════════════════════════
_DEFAULT_PAGE_CONTENT_WIDTH = Twips(12240 - 1800 - 1800)  # US-Letter minus 1" margins, python-docx default
_MIN_COL_WIDTH = Twips(300)  # floor so a single-character column doesn't collapse to nothing
_TBLW_SUCCESSORS = (
    'w:jc', 'w:tblCellSpacing', 'w:tblInd', 'w:tblBorders', 'w:shd',
    'w:tblLayout', 'w:tblCellMar', 'w:tblLook', 'w:tblCaption', 'w:tblDescription', 'w:tblPrChange',
)


_BORDER_SUCCESSORS = ('w:shd', 'w:tblLayout', 'w:tblCellMar', 'w:tblLook', 'w:tblCaption', 'w:tblDescription', 'w:tblPrChange')
_BORDER_EDGES = ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')


def _ensure_table_borders(tblPr):
    """Add plain single-line borders (all edges + inner gridlines) if the
    table doesn't already define its own <w:tblBorders>. Pandoc's default
    "Table" style has none, which OOXML renders as invisible — fine in Word's
    editing view (dashed guide lines only) but blank once exported to PDF.
    """
    from docx.oxml import OxmlElement
    if tblPr.find(qn('w:tblBorders')) is not None:
        return
    borders = OxmlElement('w:tblBorders')
    for edge in _BORDER_EDGES:
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), '4')   # eighths of a point -> 0.5pt
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), 'auto')
        borders.append(el)
    tblPr.insert_element_before(borders, *_BORDER_SUCCESSORS)


def _set_tblW_dxa(tblPr, twips):
    """Set <w:tblPr>'s <w:tblW> to an explicit dxa (twips) width, replacing
    whatever type/value pandoc put there (typically type="pct" w="0.0",
    which is otherwise meaningless once we've pinned a fixed layout).
    """
    from docx.oxml import OxmlElement
    tblW = tblPr.find(qn('w:tblW'))
    if tblW is None:
        tblW = OxmlElement('w:tblW')
        tblPr.insert_element_before(tblW, *_TBLW_SUCCESSORS)
    tblW.set(qn('w:type'), 'dxa')
    tblW.set(qn('w:w'), str(twips))


def _text_weight(text):
    """Rough content-width weight for a string: wide (CJK/fullwidth) chars
    count double a narrow char, approximating how much horizontal space
    Word's layout engine actually gives them.
    """
    weight = 0
    for ch in text:
        eaw = unicodedata.east_asian_width(ch)
        weight += 2 if eaw in ('W', 'F') else 1
    return weight


def _column_weight(table, col_idx, n_cols):
    """Max content weight across all cells landing in grid column col_idx,
    accounting for column-spanning cells (whose weight is split evenly
    across the columns they span) and vertically-merged cells (only the
    "root" cell, which python-docx already collapses to via row.cells).
    """
    best = 0
    for row in table.rows:
        cells = row.cells
        if len(cells) != n_cols:
            continue  # row has grid_before/after gaps or ragged cell count — skip for width purposes
        cell = cells[col_idx]
        span = cell.grid_span
        text = cell.text
        w = _text_weight(text) / span if text else 0
        best = max(best, w)
    return best


def autofit_tables(doc, page_content_width=None):
    """Resize every top-level table's columns to fit their content, the way
    Word's "AutoFit to Contents" does, and pin the layout so ONLYOFFICE/Word
    render exactly those widths instead of recomputing their own.

    page_content_width: usable width (a Length) to distribute across a
    table's columns; defaults to the document's own section width/margins
    if set, else a US-Letter-minus-1"-margins fallback (pandoc's generated
    docx has no explicit sectPr, so section width is usually None).

    Returns the number of tables resized. Tables with ragged rows (grid_span
    mismatches across rows) or a single column are left untouched, since a
    "proportional to content" resize wouldn't have a well-defined column
    weight there.
    """
    if page_content_width is None:
        sec = doc.sections[0] if doc.sections else None
        if sec and sec.page_width and sec.left_margin is not None and sec.right_margin is not None:
            page_content_width = sec.page_width - sec.left_margin - sec.right_margin
        else:
            page_content_width = _DEFAULT_PAGE_CONTENT_WIDTH

    resized = 0
    for table in doc.tables:
        n_cols = max((len(row.cells) for row in table.rows), default=0)
        if n_cols < 2:
            continue

        weights = [_column_weight(table, c, n_cols) for c in range(n_cols)]
        if sum(weights) <= 0:
            continue

        total_weight = sum(weights)
        col_widths = [max(_MIN_COL_WIDTH, Twips(round(page_content_width.twips * w / total_weight)))
                      for w in weights]
        # Re-normalize so the floor above doesn't push the sum past the page width.
        total_now = sum(cw.twips for cw in col_widths)
        if total_now > page_content_width.twips:
            scale = page_content_width.twips / total_now
            col_widths = [Twips(max(_MIN_COL_WIDTH.twips, round(cw.twips * scale))) for cw in col_widths]

        tbl = table._tbl
        tblPr = tbl.tblPr
        tblPr.autofit = False  # writes <w:tblLayout w:type="fixed"/>
        _set_tblW_dxa(tblPr, sum(cw.twips for cw in col_widths))
        _ensure_table_borders(tblPr)

        tblGrid = tbl.tblGrid
        tblGrid.remove_all('w:gridCol')
        for width in col_widths:
            tblGrid.add_gridCol().w = width

        for row in table.rows:
            cells = row.cells
            if len(cells) != n_cols:
                continue
            for c, cell in enumerate(cells):
                span = cell.grid_span
                cell.width = Twips(sum(cw.twips for cw in col_widths[c:c + span]))

        resized += 1

    return resized


def _resolve_cell(doc, cell_path):
    """cell_path like 'table[12].row[3].cell[5]' -> docx Cell object."""
    m = re.match(r'^table\[(\d+)\]\.row\[(\d+)\]\.cell\[(\d+)\]$', cell_path)
    if not m:
        raise ValueError(f'invalid cell path: {cell_path!r}')
    table_idx, row_idx, cell_idx = (int(g) for g in m.groups())
    items = list(doc.iter_inner_content())
    if table_idx >= len(items) or not isinstance(items[table_idx], Table):
        raise ValueError(f'no table at index {table_idx}')
    table = items[table_idx]
    if row_idx >= len(table.rows):
        raise ValueError(f'no row {row_idx} in table[{table_idx}]')
    row = table.rows[row_idx]
    if cell_idx >= len(row.cells):
        raise ValueError(f'no cell {cell_idx} in table[{table_idx}].row[{row_idx}]')
    return row.cells[cell_idx]


# ════════════════════════════════════════════════════════════════════
#  Page-number cross-reference: bookmark the source text, insert a
#  PAGEREF field in the target cell. Raw OOXML — python-docx has no
#  built-in bookmark/field API.
# ════════════════════════════════════════════════════════════════════
def add_page_crossref(doc, keyword, cell_path, paragraph_path=None):
    """Bookmark an occurrence of `keyword` in the document body, then insert
    a PAGEREF field pointing at that bookmark into the target cell (see
    list_tables for path format).

    If `paragraph_path` is given (e.g. "paragraph[5]", from list_body_paragraphs),
    the match is taken from that specific paragraph — this is how a caller
    disambiguates when `keyword` occurs more than once in the body. Without
    it, the keyword must be unique across the whole body; 0 matches raises
    "not found", 2+ matches raises "ambiguous" (callers should have the user
    pick a paragraph and pass paragraph_path instead of guessing).

    Returns the bookmark name. Raises ValueError on no-match/ambiguous-match/
    bad cell_path.

    Page numbers only become correct after the caller round-trips the saved
    document through server.py's `_recalculate_fields_docx`.
    """
    if paragraph_path is not None:
        target_paragraph, start, end = _find_keyword_in_paragraph(doc, keyword, paragraph_path)
        if target_paragraph is None:
            raise ValueError(f'keyword {keyword!r} not found in {paragraph_path}')
    else:
        matches = _find_body_keyword_all(doc, keyword)
        if not matches:
            raise ValueError(f'keyword not found in document body: {keyword!r}')
        if len(matches) > 1:
            raise ValueError(
                f'keyword {keyword!r} matches {len(matches)} paragraphs — '
                'pass paragraph_path to disambiguate which occurrence to bookmark'
            )
        target_paragraph, start, end = matches[0]

    bookmark_name = f'bm_{uuid.uuid4().hex[:12]}'
    _wrap_bookmark(target_paragraph, start, end, bookmark_name)

    cell = _resolve_cell(doc, cell_path)
    _insert_pageref_field(cell, bookmark_name)
    return bookmark_name


def _find_body_keyword_all(doc, keyword):
    """Search every body paragraph (top level only, not inside tables) for
    the first exact substring match of `keyword`. Returns a list of
    (paragraph, start, end) — one entry per matching paragraph (only the
    first occurrence within each paragraph, since one bookmark per call
    is all add_page_crossref needs).
    """
    out = []
    for item in doc.iter_inner_content():
        if not isinstance(item, Paragraph):
            continue
        pos = item.text.find(keyword)
        if pos != -1:
            out.append((item, pos, pos + len(keyword)))
    return out


def _find_keyword_in_paragraph(doc, keyword, paragraph_path):
    """Like _find_body_keyword_all, but restricted to one specific paragraph
    (identified by the "paragraph[N]" path from list_body_paragraphs).
    Returns (paragraph, start, end) or (None, None, None).
    """
    m = re.match(r'^paragraph\[(\d+)\]$', paragraph_path)
    if not m:
        raise ValueError(f'invalid paragraph path: {paragraph_path!r}')
    idx = int(m.group(1))
    items = list(doc.iter_inner_content())
    if idx >= len(items) or not isinstance(items[idx], Paragraph):
        raise ValueError(f'no paragraph at index {idx}')
    paragraph = items[idx]
    pos = paragraph.text.find(keyword)
    if pos == -1:
        return None, None, None
    return paragraph, pos, pos + len(keyword)


def _wrap_bookmark(paragraph, start, end, bookmark_name):
    """Insert <w:bookmarkStart>/<w:bookmarkEnd> around paragraph.text[start:end]
    by splitting whichever run(s) the span touches, so the bookmark markers
    sit exactly at the span boundaries.
    """
    bookmark_id = str(abs(hash(bookmark_name)) % 1000000)

    start_elem = _make_bookmark_elem('w:bookmarkStart', bookmark_id, bookmark_name)
    end_elem = _make_bookmark_elem('w:bookmarkEnd', bookmark_id, None)

    # Split the run containing `start` (if any) so the bookmark start lands
    # cleanly between runs, then likewise for `end`.
    _split_run_at(paragraph, start)
    _split_run_at(paragraph, end)

    spans = _run_spans(paragraph)  # recompute after splits
    start_run_elem = None
    end_run_elem = None
    for run, r_start, r_end in spans:
        if r_start == start:
            start_run_elem = run._r
        if r_end == end:
            end_run_elem = run._r

    p_elem = paragraph._p
    if start_run_elem is not None:
        start_run_elem.addprevious(start_elem)
    else:
        p_elem.append(start_elem)
    if end_run_elem is not None:
        end_run_elem.addnext(end_elem)
    else:
        p_elem.append(end_elem)


def _split_run_at(paragraph, char_pos):
    """Ensure a run boundary exists at `char_pos` in paragraph.text, splitting
    a run into two runs (same formatting) if char_pos falls in its middle.
    No-op if char_pos is already a boundary or out of range.
    """
    if char_pos <= 0 or char_pos >= len(paragraph.text):
        return
    spans = _run_spans(paragraph)
    for run, r_start, r_end in spans:
        if r_start < char_pos < r_end:
            local = char_pos - r_start
            full_text = run.text
            run.text = full_text[:local]
            new_r = _clone_run_element(run._r)
            _set_run_text(new_r, full_text[local:])
            run._r.addnext(new_r)
            return


def _clone_run_element(run_elem):
    import copy
    return copy.deepcopy(run_elem)


def _set_run_text(run_elem, text):
    t_elem = run_elem.find(qn('w:t'))
    if t_elem is None:
        t_elem = run_elem.makeelement(qn('w:t'), {})
        run_elem.append(t_elem)
    t_elem.text = text
    t_elem.set(qn('xml:space'), 'preserve')


def _make_bookmark_elem(tag, bookmark_id, name):
    from docx.oxml import OxmlElement
    elem = OxmlElement(tag)
    elem.set(qn('w:id'), bookmark_id)
    if name is not None:
        elem.set(qn('w:name'), name)
    return elem


def _insert_pageref_field(cell, bookmark_name):
    """Insert a PAGEREF field (raw fldChar/instrText/fldChar sequence) into
    the target cell's first paragraph, referencing bookmark_name.

    If the cell already contains a "第...页" placeholder (the common pattern
    in index tables, e.g. "第　　页" with blank fill-in space), the field is
    spliced into that gap so the result reads "第4页" — the surrounding
    wording is preserved. Otherwise the field is appended to the existing
    text wrapped in "第...页", so nothing already in the cell is discarded.

    The displayed value is a placeholder ("1") until ONLYOFFICE recalculates
    it — see server.py `_recalculate_fields_docx`.
    """
    from docx.oxml import OxmlElement
    paragraph = cell.paragraphs[0]
    text = paragraph.text
    m = re.search(r'第[\s　]*页', text)
    if m:
        prefix, suffix = text[:m.start()], text[m.end():]
    else:
        prefix, suffix = text, ''

    for run in list(paragraph.runs):
        run._r.getparent().remove(run._r)
    p_elem = paragraph._p

    def make_run(children=None, text_content=None):
        r = OxmlElement('w:r')
        if text_content is not None:
            t = OxmlElement('w:t')
            t.set(qn('xml:space'), 'preserve')
            t.text = text_content
            r.append(t)
        for child in (children or []):
            r.append(child)
        return r

    if prefix:
        p_elem.append(make_run(text_content=prefix))
    p_elem.append(make_run(text_content='第'))

    fld_begin = OxmlElement('w:fldChar')
    fld_begin.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = f' PAGEREF {bookmark_name} \\h '
    fld_separate = OxmlElement('w:fldChar')
    fld_separate.set(qn('w:fldCharType'), 'separate')
    placeholder_text = OxmlElement('w:t')
    placeholder_text.text = '1'
    fld_end = OxmlElement('w:fldChar')
    fld_end.set(qn('w:fldCharType'), 'end')

    p_elem.append(make_run([fld_begin]))
    p_elem.append(make_run([instr]))
    p_elem.append(make_run([fld_separate]))
    p_elem.append(make_run([placeholder_text]))
    p_elem.append(make_run([fld_end]))

    p_elem.append(make_run(text_content='页'))
    if suffix:
        p_elem.append(make_run(text_content=suffix))
