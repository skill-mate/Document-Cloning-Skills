#!/usr/bin/env python3
"""
build_from_skeleton.py — Use a reference .docx as a SKELETON: preserve ALL of its
styles.xml, numbering.xml, theme/, header*.xml, footer*.xml, settings.xml, and
section properties (page setup, header/footer references). Replace ONLY the body
content with material from Markdown.

This is the most precise way to produce "the same document with new content":
every visual decision the reference made (fonts, sizes, colours, numbering,
header banner, page numbers, page setup) is inherited byte-for-byte.

Usage:
    build_from_skeleton.py INPUT.md --reference REF.docx -o OUT.docx
        [--no-h1-pagebreak]

Strategy:
  1. Open REF.docx with python-docx.
  2. Snapshot the trailing <w:sectPr>; remove every <w:p>/<w:tbl>/etc from <w:body>;
     re-attach the sectPr so the body is empty but section properties survive.
  3. Render Markdown blocks using styles BY NAME — Heading 1, 标题 1, Title,
     副标题, Caption, List Number, Quote, Table Grid, etc — first match wins.
     The actual visual style comes from the reference's styles.xml.
  4. For structural guarantees we still touch XML directly:
       - cover section gets w:vAlign=center
       - toc gets a TOC field
       - H1 paragraphs get w:pageBreakBefore
       - every table row gets w:cantSplit; first row gets w:tblHeader
"""
from __future__ import annotations
import argparse, re, sys
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from build_docx import parse_md  # only the parser is reused


# ---------- skeleton bootstrap ----------

def _clear_body_keep_sectpr(doc):
    """Remove every direct child of <w:body> EXCEPT the trailing <w:sectPr>.
    The sectPr carries page setup + header/footer references and must be
    preserved so all subsequent sections (added via add_section) inherit it."""
    body = doc.element.body
    final_sectPr_orig = body.find(qn("w:sectPr"))
    final_sectPr = deepcopy(final_sectPr_orig) if final_sectPr_orig is not None else None
    for child in list(body):
        body.remove(child)
    if final_sectPr is not None:
        body.append(final_sectPr)


# ---------- style discovery ----------

def _find_style(doc, *candidates):
    """Return the first Style OBJECT whose name matches any candidate, else None.
    We must return the object (not the name) because doc.styles[name] lookup is
    flaky on documents whose style_id differs from name (e.g. id='Heading2'
    name='Heading 2' — the indexer keys on style_id and raises KeyError when
    you pass the name)."""
    by_name = {}
    for s in doc.styles:
        try:
            n = s.name
        except Exception:
            continue
        if n and n not in by_name:
            by_name[n] = s
    for name in candidates:
        if name in by_name:
            return by_name[name]
    return None


def _heading_style(doc, level):
    return _find_style(doc, f"Heading {level}", f"标题 {level}", f"Heading{level}")


def _table_style(doc):
    return _find_style(
        doc,
        "Table Grid", "网格型", "Light Grid", "Light Shading",
        "Medium Shading 1", "Plain Table 1", "Table Grid Light",
    )


def _apply_style(p, style_obj):
    """Apply a Style object to a paragraph. No-op if None."""
    if style_obj is not None:
        p.style = style_obj
    return p


# ---------- inline runs (let paragraph style decide fonts) ----------

INLINE_RE = re.compile(r"(\*\*([^*]+?)\*\*|\*([^*\n]+?)\*|`([^`]+?)`)")

def _add_inline(p, text):
    """Add runs to paragraph p. Plain runs inherit the paragraph's style
    (which inherits from the reference's styles.xml). Only toggle bold/italic
    or set Consolas for `code` spans."""
    pos = 0
    for m in INLINE_RE.finditer(text):
        if m.start() > pos:
            p.add_run(text[pos:m.start()])
        if m.group(2) is not None:
            r = p.add_run(m.group(2)); r.bold = True
        elif m.group(3) is not None:
            r = p.add_run(m.group(3)); r.italic = True
        elif m.group(4) is not None:
            r = p.add_run(m.group(4)); r.font.name = "Consolas"
        pos = m.end()
    if pos < len(text):
        p.add_run(text[pos:])


# ---------- low-level XML helpers ----------

def _para_pagebreak_before(p):
    pPr = p._p.get_or_add_pPr()
    pPr.append(OxmlElement("w:pageBreakBefore"))


def _add_pagebreak(doc):
    p = doc.add_paragraph(); r = p.add_run()
    br = OxmlElement("w:br"); br.set(qn("w:type"), "page")
    r._r.append(br)


def _new_section(doc):
    return doc.add_section(WD_SECTION.NEW_PAGE)


def _set_section_vertical_center(section):
    sectPr = section._sectPr
    vAlign = sectPr.find(qn("w:vAlign"))
    if vAlign is None:
        vAlign = OxmlElement("w:vAlign"); sectPr.append(vAlign)
    vAlign.set(qn("w:val"), "center")


def _clear_section_vertical_align(section):
    sectPr = section._sectPr
    vAlign = sectPr.find(qn("w:vAlign"))
    if vAlign is not None:
        sectPr.remove(vAlign)


# ---------- block renderers ----------

def render_cover(doc, fields):
    title = fields.get("title", "")
    subtitle = fields.get("subtitle", "")
    org = fields.get("org", "")
    date = fields.get("date", "")

    _set_section_vertical_center(doc.sections[-1])

    title_style = _find_style(doc, "Title", "标题", "Heading 1", "标题 1")
    if title:
        p = doc.add_paragraph()
        _apply_style(p, title_style)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _add_inline(p, title)
    if subtitle:
        sub_style = _find_style(doc, "Subtitle", "副标题", "Heading 2")
        p = doc.add_paragraph()
        _apply_style(p, sub_style)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _add_inline(p, subtitle)
    for _ in range(6):
        doc.add_paragraph()
    if org:
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(org); r.bold = True
    if date:
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run(date)

    _new_section(doc)
    _clear_section_vertical_align(doc.sections[-1])


def render_toc(doc, depth):
    title_style = _find_style(doc, "TOC Heading", "Heading 1", "标题 1")
    p = doc.add_paragraph()
    _apply_style(p, title_style)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("目  录")

    p = doc.add_paragraph()
    r = p.add_run()
    fc1 = OxmlElement("w:fldChar"); fc1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve")
    instr.text = f' TOC \\o "1-{depth}" \\h \\z \\u '
    fc2 = OxmlElement("w:fldChar"); fc2.set(qn("w:fldCharType"), "separate")
    placeholder = OxmlElement("w:t"); placeholder.text = "右键此处选择「更新域」以生成目录"
    fc4 = OxmlElement("w:fldChar"); fc4.set(qn("w:fldCharType"), "end")
    for el in (fc1, instr, fc2, placeholder, fc4):
        r._r.append(el)

    _new_section(doc)
    _clear_section_vertical_align(doc.sections[-1])


def render_heading(doc, level, text, auto_h1_pb=True):
    style_obj = _heading_style(doc, min(level, 4))
    p = doc.add_paragraph()
    _apply_style(p, style_obj)
    if level == 1 and auto_h1_pb:
        _para_pagebreak_before(p)
    _add_inline(p, text)


def render_para(doc, text):
    p = doc.add_paragraph()
    _add_inline(p, text)


def render_quote(doc, text):
    style = _find_style(doc, "Quote", "引用", "Intense Quote")
    p = doc.add_paragraph()
    _apply_style(p, style)
    _add_inline(p, text)


def render_list(doc, items, ordered):
    if ordered:
        style = _find_style(doc, "List Number", "List Paragraph")
    else:
        style = _find_style(doc, "List Bullet", "List Paragraph")
    for idx, item in enumerate(items, 1):
        p = doc.add_paragraph()
        _apply_style(p, style)
        if style is None:
            p.add_run(f"{idx}. " if ordered else "• ")
        _add_inline(p, item)


def render_code(doc, text):
    style = _find_style(doc, "HTML Preformatted", "Plain Text")
    p = doc.add_paragraph()
    _apply_style(p, style)
    r = p.add_run(text); r.font.name = "Consolas"


def render_image(doc, src, alt, caption):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if Path(src).exists():
        try:
            p.add_run().add_picture(src, width=Cm(14))
        except Exception as e:
            p.add_run(f"[图片：{src} 加载失败 {e}]")
    else:
        p.add_run(f"[图片占位：{alt or src}]")
    if caption:
        cstyle = _find_style(doc, "Caption", "题注")
        cp = doc.add_paragraph()
        _apply_style(cp, cstyle)
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cp.add_run(caption)


def _add_table_borders(table):
    """Manually add single-line black borders to every cell — used when the
    reference doesn't ship a Table Grid style."""
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr"); tbl.insert(0, tblPr)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{edge}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "4")
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "000000")
        borders.append(b)
    tblPr.append(borders)


def render_table(doc, header, rows, caption, widths):
    if caption:
        cstyle = _find_style(doc, "Caption", "题注")
        cp = doc.add_paragraph()
        _apply_style(cp, cstyle)
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cp.add_run(caption)

    cols = len(header)
    table = doc.add_table(rows=1 + len(rows), cols=cols)
    style_obj = _table_style(doc)
    if style_obj is not None:
        try:
            table.style = style_obj
        except Exception:
            _add_table_borders(table)
    else:
        # reference has no table style → add borders manually
        _add_table_borders(table)

    # fixed layout
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr"); tbl.insert(0, tblPr)
    layout = OxmlElement("w:tblLayout"); layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    table.autofit = False

    # widths
    if widths and len(widths) == cols:
        sec = doc.sections[-1]
        usable = sec.page_width - sec.left_margin - sec.right_margin
        total = sum(widths)
        for c in range(cols):
            w = int(usable * widths[c] / total)
            for row in table.rows:
                row.cells[c].width = w

    # header row
    for c, txt in enumerate(header):
        cell = table.rows[0].cells[c]
        para = cell.paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.first_line_indent = Pt(0)
        r = para.add_run(txt); r.bold = True
    # body rows
    for ri, row in enumerate(rows, start=1):
        for c, txt in enumerate(row[:cols]):
            cell = table.rows[ri].cells[c]
            para = cell.paragraphs[0]
            para.paragraph_format.first_line_indent = Pt(0)
            parts = str(txt).split("<br>") if "<br>" in str(txt) else [str(txt)]
            for pi, ptxt in enumerate(parts):
                if pi > 0:
                    para = cell.add_paragraph()
                    para.paragraph_format.first_line_indent = Pt(0)
                _add_inline(para, ptxt)

    # cantSplit on every row
    for row in table.rows:
        trPr = row._tr.get_or_add_trPr()
        trPr.append(OxmlElement("w:cantSplit"))
    # tblHeader on first row
    trPr = table.rows[0]._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:tblHeader"))


# ---------- main render driver ----------

def render(blocks, doc, *, auto_h1_pagebreak=True):
    for b in blocks:
        k = b["kind"]
        if k == "cover":
            render_cover(doc, b["fields"])
        elif k == "toc":
            render_toc(doc, b["depth"])
        elif k == "pagebreak":
            _add_pagebreak(doc)
        elif k == "newsection":
            _new_section(doc)
            _clear_section_vertical_align(doc.sections[-1])
        elif k == "heading":
            render_heading(doc, b["level"], b["text"], auto_h1_pb=auto_h1_pagebreak)
        elif k == "para":
            render_para(doc, b["text"])
        elif k == "quote":
            render_quote(doc, b["text"])
        elif k == "list":
            render_list(doc, b["items"], b["ordered"])
        elif k == "code":
            render_code(doc, b["text"])
        elif k == "image":
            render_image(doc, b["src"], b["alt"], b["caption"])
        elif k == "table":
            render_table(doc, b["header"], b["rows"], b["caption"], b["widths"])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input_md")
    ap.add_argument("--reference", required=True, help="reference .docx (skeleton)")
    ap.add_argument("-o", "--output", required=True)
    ap.add_argument("--no-h1-pagebreak", action="store_true",
                    help="don't auto-page-break before each H1 (default: do)")
    args = ap.parse_args()

    doc = Document(args.reference)
    _clear_body_keep_sectpr(doc)
    blocks = parse_md(Path(args.input_md).read_text(encoding="utf-8"))
    render(blocks, doc, auto_h1_pagebreak=not args.no_h1_pagebreak)
    doc.save(args.output)

    # self-check
    Document(args.output)
    print(f"OK: wrote {args.output} ({Path(args.output).stat().st_size} bytes)")
    print(f"  reference styles: {sum(1 for _ in doc.styles)}")
    used_h = []
    for n in (1, 2, 3, 4):
        s = _heading_style(doc, n)
        if s is not None:
            used_h.append(s.name)
    print(f"  heading styles bound: {', '.join(used_h) or '(none, fell back to Normal)'}")
    ts = _table_style(doc)
    print(f"  table style bound: {ts.name if ts else '(default Table Grid)'}")


if __name__ == "__main__":
    main()
