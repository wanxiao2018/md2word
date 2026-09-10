"""Apply academic Word body styles and strip section-divider rules."""

from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
OMATH_PARA = f"{{{MATH_NS}}}oMathPara"
OMATH = f"{{{MATH_NS}}}oMath"

WESTERN_FONT = "Times New Roman"
EAST_ASIA_BODY = "宋体"
EAST_ASIA_HEADING = "黑体"
CODE_FONTS = {"Consolas", "Courier New", "Courier", "Cambria Math"}
CODE_STYLES = {"SourceCode", "VerbatimChar", "Token"}

BODY_STYLE_IDS = {"Normal", "BodyText", "FirstParagraph"}
HEADING_STYLE_IDS = {f"Heading{i}" for i in range(1, 10)}
NO_INDENT_STYLE_IDS = {
    "Compact",
    "ListParagraph",
    "SourceCode",
    "DefinitionTerm",
    "Definition",
    "BlockText",
    "Quote",
    "IntenseQuote",
    "Caption",
    "TableCaption",
    "ImageCaption",
    "Title",
    "Subtitle",
    "TOCHeading",
}
CODE_STYLE_IDS = {"SourceCode"}
CAPTION_STYLE_IDS = {"Caption", "TableCaption", "ImageCaption"}
HEADING_SIZES = {
    "Heading1": 18,
    "Heading2": 16,
    "Heading3": 14,
    "Heading4": 13,
    "Heading5": 12,
    "Heading6": 12,
}
CAPTION_RE = re.compile(r"^(表|表格|Table)\s*\.?\s*\d+", re.I)
REF_HEADING_RE = re.compile(r"^(参考文献|参考书目|References|Bibliography)\s*$", re.I)
THEME_FONT_ATTRS = (
    "w:asciiTheme",
    "w:hAnsiTheme",
    "w:cstheme",
    "w:eastAsiaTheme",
    "w:hint",
)


def _get_or_add(parent, tag: str):
    el = parent.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        parent.append(el)
    return el


def _style_by_id(doc: Document, style_id: str):
    for style in doc.styles:
        if getattr(style, "style_id", None) == style_id:
            return style
    return None


def _clear_theme_fonts(rFonts) -> None:
    for attr in THEME_FONT_ATTRS:
        key = qn(attr)
        if key in rFonts.attrib:
            del rFonts.attrib[key]


def _set_run_fonts(style, *, western: str, east_asia: str, size_pt: int, bold=None, color=None) -> None:
    style.font.name = western
    style.font.size = Pt(size_pt)
    if bold is not None:
        style.font.bold = bold
    if color is not None:
        style.font.color.rgb = color
    rPr = style.element.get_or_add_rPr()
    rFonts = _get_or_add(rPr, "w:rFonts")
    rFonts.set(qn("w:ascii"), western)
    rFonts.set(qn("w:hAnsi"), western)
    rFonts.set(qn("w:eastAsia"), east_asia)
    rFonts.set(qn("w:cs"), western)
    _clear_theme_fonts(rFonts)
    sz = _get_or_add(rPr, "w:sz")
    sz.set(qn("w:val"), str(size_pt * 2))
    szCs = _get_or_add(rPr, "w:szCs")
    szCs.set(qn("w:val"), str(size_pt * 2))


def _set_paragraph_format(
    style,
    *,
    first_line_chars: float | None = None,
    align: str | None = None,
    line_twips: int | None = None,
    before: int = 0,
    after: int = 0,
    drop_border: bool = True,
) -> None:
    pPr = style.element.get_or_add_pPr()
    if drop_border:
        pBdr = pPr.find(qn("w:pBdr"))
        if pBdr is not None:
            pPr.remove(pBdr)

    spacing = _get_or_add(pPr, "w:spacing")
    spacing.set(qn("w:before"), str(before))
    spacing.set(qn("w:after"), str(after))
    if before == 0:
        spacing.set(qn("w:beforeLines"), "0")
    if after == 0:
        spacing.set(qn("w:afterLines"), "0")
    if line_twips is not None:
        spacing.set(qn("w:line"), str(line_twips))
        spacing.set(qn("w:lineRule"), "auto")

    if first_line_chars is not None:
        _set_indent(pPr, first_line_chars=first_line_chars)

    if align:
        jc = _get_or_add(pPr, "w:jc")
        jc.set(qn("w:val"), align)


def _remove_paragraph_borders(doc: Document) -> None:
    body = doc.element.body
    for pBdr in list(body.iter(qn("w:pBdr"))):
        parent = pBdr.getparent()
        if parent is not None:
            parent.remove(pBdr)


def _is_in_table(p_el) -> bool:
    parent = p_el.getparent()
    while parent is not None:
        if parent.tag == qn("w:tbl"):
            return True
        parent = parent.getparent()
    return False


def _is_inside_math(el) -> bool:
    parent = el.getparent()
    while parent is not None:
        if parent.tag in {OMATH, OMATH_PARA}:
            return True
        parent = parent.getparent()
    return False


def _para_style_id(p_el) -> str:
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        return ""
    pStyle = pPr.find(qn("w:pStyle"))
    if pStyle is None:
        return ""
    return pStyle.get(qn("w:val")) or ""


def _para_is_centered(p_el) -> bool:
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        return False
    jc = pPr.find(qn("w:jc"))
    if jc is None:
        return False
    return (jc.get(qn("w:val")) or "") == "center"


def _para_has_display_math(p_el) -> bool:
    for child in p_el.iter():
        if child.tag == OMATH_PARA:
            return True
    return False


def _para_has_num(p_el) -> bool:
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        return False
    return pPr.find(qn("w:numPr")) is not None


def _para_text(p_el) -> str:
    return "".join(t.text or "" for t in p_el.iter(qn("w:t")))


def _clear_first_line_indent(pPr) -> None:
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        return
    for attr in (qn("w:firstLine"), qn("w:firstLineChars")):
        if attr in ind.attrib:
            del ind.attrib[attr]
    if not ind.attrib:
        pPr.remove(ind)


def _set_indent(pPr, *, first_line_chars: float = 0, left_chars: float = 0, right_chars: float = 0) -> None:
    """Match Word's 字符 indent fields. 2 chars ≈ 480 twips at 12pt."""
    ind = _get_or_add(pPr, "w:ind")
    if qn("w:hanging") in ind.attrib:
        del ind.attrib[qn("w:hanging")]
    ind.set(qn("w:leftChars"), str(int(round(left_chars * 100))))
    ind.set(qn("w:left"), "0" if left_chars == 0 else str(int(round(left_chars * 240))))
    ind.set(qn("w:rightChars"), str(int(round(right_chars * 100))))
    ind.set(qn("w:right"), "0" if right_chars == 0 else str(int(round(right_chars * 240))))
    if first_line_chars:
        ind.set(qn("w:firstLineChars"), str(int(round(first_line_chars * 100))))
        ind.set(qn("w:firstLine"), "480" if first_line_chars == 2 else str(int(round(first_line_chars * 240))))
    else:
        ind.set(qn("w:firstLineChars"), "0")
        ind.set(qn("w:firstLine"), "0")


def _set_jc(pPr, align: str) -> None:
    jc = _get_or_add(pPr, "w:jc")
    jc.set(qn("w:val"), align)


def _set_spacing(pPr, *, before: str | None = None, after: str | None = None, line: str | None = None) -> None:
    spacing = _get_or_add(pPr, "w:spacing")
    if before is not None:
        spacing.set(qn("w:before"), before)
        if before == "0":
            spacing.set(qn("w:beforeLines"), "0")
    if after is not None:
        spacing.set(qn("w:after"), after)
        if after == "0":
            spacing.set(qn("w:afterLines"), "0")
    if line is not None:
        spacing.set(qn("w:line"), line)
        spacing.set(qn("w:lineRule"), "auto")


def _set_first_line(pPr, chars: int = 200, twips: str = "480") -> None:
    del twips  # kept for call-site compatibility; indent is character-based
    _set_indent(pPr, first_line_chars=chars / 100)


def _looks_like_caption(p_el) -> bool:
    nxt = p_el.getnext()
    if nxt is None or nxt.tag != qn("w:tbl"):
        return False
    return bool(CAPTION_RE.match(_para_text(p_el).strip()))


def _is_heading_para(sid: str) -> bool:
    return sid in HEADING_STYLE_IDS or sid.startswith("Heading")


def _run_is_code(r) -> bool:
    rPr = r.find(qn("w:rPr"))
    if rPr is None:
        return False
    rStyle = rPr.find(qn("w:rStyle"))
    if rStyle is not None and (rStyle.get(qn("w:val")) or "") in CODE_STYLES:
        return True
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        return False
    ascii_font = rFonts.get(qn("w:ascii")) or ""
    return ascii_font in CODE_FONTS


def _apply_run_fonts(r, *, western: str, east_asia: str) -> None:
    if _is_inside_math(r) or _run_is_code(r):
        return
    rPr = _get_or_add(r, "w:rPr")
    rFonts = _get_or_add(rPr, "w:rFonts")
    rFonts.set(qn("w:ascii"), western)
    rFonts.set(qn("w:hAnsi"), western)
    rFonts.set(qn("w:cs"), western)
    rFonts.set(qn("w:eastAsia"), east_asia)
    _clear_theme_fonts(rFonts)


def _force_para_fonts(p_el, *, heading: bool = False) -> None:
    east = EAST_ASIA_HEADING if heading else EAST_ASIA_BODY
    for r in p_el.iter(qn("w:r")):
        _apply_run_fonts(r, western=WESTERN_FONT, east_asia=east)


def _set_border(parent, tag: str, *, sz: str = "4", color: str = "666666") -> None:
    el = _get_or_add(parent, tag)
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), sz)
    el.set(qn("w:space"), "0")
    el.set(qn("w:color"), color)


def _style_tables(doc: Document) -> None:
    body = doc.element.body
    for tbl in body.iter(qn("w:tbl")):
        tblPr = _get_or_add(tbl, "w:tblPr")
        tblW = _get_or_add(tblPr, "w:tblW")
        tblW.set(qn("w:w"), "5000")
        tblW.set(qn("w:type"), "pct")
        jc = _get_or_add(tblPr, "w:jc")
        jc.set(qn("w:val"), "center")
        layout = _get_or_add(tblPr, "w:tblLayout")
        layout.set(qn("w:type"), "autofit")

        borders = _get_or_add(tblPr, "w:tblBorders")
        for edge in ("w:top", "w:left", "w:bottom", "w:right", "w:insideH", "w:insideV"):
            _set_border(borders, edge)

        cell_mar = _get_or_add(tblPr, "w:tblCellMar")
        for side, val in (("w:top", "60"), ("w:left", "80"), ("w:bottom", "60"), ("w:right", "80")):
            el = _get_or_add(cell_mar, side)
            el.set(qn("w:w"), val)
            el.set(qn("w:type"), "dxa")

        rows = list(tbl.findall(qn("w:tr")))
        for ridx, tr in enumerate(rows):
            for tc in tr.findall(qn("w:tc")):
                tcPr = _get_or_add(tc, "w:tcPr")
                vAlign = _get_or_add(tcPr, "w:vAlign")
                vAlign.set(qn("w:val"), "center")
                tc_borders = _get_or_add(tcPr, "w:tcBorders")
                for edge in ("w:top", "w:left", "w:bottom", "w:right"):
                    _set_border(tc_borders, edge)
                if ridx == 0:
                    shd = _get_or_add(tcPr, "w:shd")
                    shd.set(qn("w:val"), "clear")
                    shd.set(qn("w:color"), "auto")
                    shd.set(qn("w:fill"), "F2F2F2")
                for p in tc.findall(qn("w:p")):
                    pPr = _get_or_add(p, "w:pPr")
                    _clear_first_line_indent(pPr)
                    if pPr.find(qn("w:jc")) is None:
                        _set_jc(pPr, "center")
                    _set_spacing(pPr, before="40", after="40", line="276")
                    _force_para_fonts(p)
                    if ridx == 0:
                        rPr = _get_or_add(pPr, "w:rPr")
                        b = _get_or_add(rPr, "w:b")
                        b.set(qn("w:val"), "1")
                        for r in p.findall(qn("w:r")):
                            rrPr = _get_or_add(r, "w:rPr")
                            _get_or_add(rrPr, "w:b")


def _polish_numbering(doc: Document) -> None:
    numbering_part = getattr(doc.part, "numbering_part", None)
    if numbering_part is None:
        return
    root = numbering_part._element
    for lvl in root.iter(qn("w:lvl")):
        pPr = _get_or_add(lvl, "w:pPr")
        _set_indent(pPr, first_line_chars=2.0)
        _set_spacing(pPr, before="0", after="0", line="360")


def _style_list_para(p_el) -> None:
    pPr = _get_or_add(p_el, "w:pPr")
    _set_indent(pPr, first_line_chars=2.0)
    _set_jc(pPr, "both")
    _set_spacing(pPr, before="0", after="0", line="360")


def _style_caption(p_el) -> None:
    pPr = _get_or_add(p_el, "w:pPr")
    _clear_first_line_indent(pPr)
    ind = _get_or_add(pPr, "w:ind")
    ind.set(qn("w:firstLineChars"), "0")
    ind.set(qn("w:firstLine"), "0")
    _set_jc(pPr, "center")
    _set_spacing(pPr, before="160", after="80", line="276")


def _style_display_math(p_el) -> None:
    pPr = _get_or_add(p_el, "w:pPr")
    _clear_first_line_indent(pPr)
    ind = _get_or_add(pPr, "w:ind")
    ind.set(qn("w:firstLineChars"), "0")
    ind.set(qn("w:firstLine"), "0")
    _set_jc(pPr, "center")
    _set_spacing(pPr, before="120", after="120", line="360")


def _style_body(p_el) -> None:
    pPr = _get_or_add(p_el, "w:pPr")
    _set_indent(pPr, first_line_chars=2.0)
    _set_jc(pPr, "both")
    _set_spacing(pPr, before="0", after="0", line="360")


def _demote_split_list_fragment(p_el) -> None:
    """Drop false numbering that Pandoc emits after display math in a list item."""
    pPr = _get_or_add(p_el, "w:pPr")
    numPr = pPr.find(qn("w:numPr"))
    if numPr is not None:
        pPr.remove(numPr)
    _set_indent(pPr, first_line_chars=2.0)
    _set_jc(pPr, "both")
    _set_spacing(pPr, before="0", after="0", line="360")


def _fix_split_list_after_math(body) -> None:
    prev = None
    for el in list(body):
        if el.tag == qn("w:tbl"):
            prev = None
            continue
        if el.tag != qn("w:p"):
            continue
        if _is_heading_para(_para_style_id(el)):
            prev = None
            continue
        math = _para_has_display_math(el)
        numbered = _para_has_num(el)
        empty = not _para_text(el).strip() and not math and not numbered
        if empty:
            continue
        if math:
            if prev in {"list", "math-after-list", "cont"}:
                prev = "math-after-list"
            else:
                prev = None
            continue
        if numbered and prev == "math-after-list":
            _demote_split_list_fragment(el)
            prev = "cont"
            continue
        if numbered:
            prev = "list"
            continue
        prev = None


def _style_reference_block(body) -> None:
    in_refs = False
    for child in list(body):
        if child.tag != qn("w:p"):
            if child.tag == qn("w:tbl") and in_refs:
                in_refs = False
            continue
        sid = _para_style_id(child)
        if _is_heading_para(sid):
            in_refs = bool(REF_HEADING_RE.match(_para_text(child).strip()))
            continue
        if not in_refs:
            continue
        pPr = _get_or_add(child, "w:pPr")
        _set_indent(pPr, first_line_chars=2.0)
        _set_jc(pPr, "both")
        _set_spacing(pPr, before="0", after="0", line="360")


def polish_docx(path: Path) -> None:
    """Force academic body layout and drop leftover section rules."""
    path = Path(path)
    doc = Document(str(path))

    normal = _style_by_id(doc, "Normal") or doc.styles["Normal"]
    _set_run_fonts(normal, western=WESTERN_FONT, east_asia=EAST_ASIA_BODY, size_pt=12)
    _set_paragraph_format(
        normal,
        first_line_chars=2.0,
        align="both",
        line_twips=360,
        before=0,
        after=0,
    )

    for sid in BODY_STYLE_IDS:
        style = _style_by_id(doc, sid)
        if style is None:
            continue
        _set_run_fonts(style, western=WESTERN_FONT, east_asia=EAST_ASIA_BODY, size_pt=12)
        _set_paragraph_format(
            style,
            first_line_chars=2.0,
            align="both",
            line_twips=360,
            before=0,
            after=0,
        )

    for sid in NO_INDENT_STYLE_IDS:
        style = _style_by_id(doc, sid)
        if style is None:
            continue
        align = "left" if sid in CODE_STYLE_IDS else "both"
        if sid in CAPTION_STYLE_IDS:
            align = "center"
        list_like = sid in {"Compact", "ListParagraph"}
        _set_paragraph_format(
            style,
            first_line_chars=2.0 if list_like else 0,
            align=align,
            line_twips=360 if list_like else 276,
            before=0 if list_like else 40,
            after=0 if list_like else 40,
        )

    for sid, size in HEADING_SIZES.items():
        style = _style_by_id(doc, sid)
        if style is None:
            continue
        _set_run_fonts(
            style,
            western=WESTERN_FONT,
            east_asia=EAST_ASIA_HEADING,
            size_pt=size,
            bold=True,
            color=RGBColor(0, 0, 0),
        )
        before = {18: 280, 16: 240, 14: 200}.get(size, 160)
        _set_paragraph_format(
            style,
            first_line_chars=0,
            align="left",
            line_twips=360,
            before=before,
            after=80,
        )

    _remove_paragraph_borders(doc)
    _polish_numbering(doc)
    _style_tables(doc)

    for p in list(doc.element.body.iter(qn("w:p"))):
        if _is_in_table(p):
            continue
        sid = _para_style_id(p)
        heading = _is_heading_para(sid)
        _force_para_fonts(p, heading=heading)
        if heading:
            continue
        if sid in CODE_STYLE_IDS:
            continue
        if _looks_like_caption(p) or sid in CAPTION_STYLE_IDS:
            _style_caption(p)
            continue
        if _para_has_display_math(p):
            _style_display_math(p)
            continue
        if _para_has_num(p) or sid in {"Compact", "ListParagraph"}:
            _style_list_para(p)
            continue
        if sid in NO_INDENT_STYLE_IDS:
            pPr = _get_or_add(p, "w:pPr")
            _clear_first_line_indent(pPr)
            if sid not in CAPTION_STYLE_IDS:
                _set_jc(pPr, "both")
            continue
        if _para_is_centered(p):
            continue
        if sid in BODY_STYLE_IDS or sid == "":
            _style_body(p)

    _fix_split_list_after_math(doc.element.body)
    _style_reference_block(doc.element.body)
    doc.save(str(path))
