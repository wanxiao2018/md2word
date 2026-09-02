"""Apply academic Word body styles and strip section-divider rules."""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
OMATH_PARA = f"{{{MATH_NS}}}oMathPara"


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
HEADING_SIZES = {
    "Heading1": 18,
    "Heading2": 16,
    "Heading3": 14,
    "Heading4": 13,
    "Heading5": 12,
    "Heading6": 12,
}


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
    if line_twips is not None:
        spacing.set(qn("w:line"), str(line_twips))
        spacing.set(qn("w:lineRule"), "auto")

    if first_line_chars is not None:
        ind = _get_or_add(pPr, "w:ind")
        if first_line_chars:
            # 200 = 2.00 characters; 480 twips ≈ 2 chars at 12pt CJK.
            ind.set(qn("w:firstLineChars"), str(int(round(first_line_chars * 100))))
            ind.set(qn("w:firstLine"), "480")
        else:
            ind.set(qn("w:firstLineChars"), "0")
            ind.set(qn("w:firstLine"), "0")
            if ind.get(qn("w:firstLineChars")) == "0":
                pass

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


def polish_docx(path: Path) -> None:
    """Force academic body layout and drop leftover section rules."""
    path = Path(path)
    doc = Document(str(path))

    normal = _style_by_id(doc, "Normal") or doc.styles["Normal"]
    _set_run_fonts(normal, western="Times New Roman", east_asia="宋体", size_pt=12)
    _set_paragraph_format(
        normal,
        first_line_chars=2.0,
        align="both",
        line_twips=360,  # 1.5 line spacing (240 = single)
        before=0,
        after=0,
    )

    for sid in BODY_STYLE_IDS:
        style = _style_by_id(doc, sid)
        if style is None:
            continue
        _set_run_fonts(style, western="Times New Roman", east_asia="宋体", size_pt=12)
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
        _set_paragraph_format(
            style,
            first_line_chars=0,
            align="left",
            line_twips=276,
            before=40,
            after=40,
        )

    for sid, size in HEADING_SIZES.items():
        style = _style_by_id(doc, sid)
        if style is None:
            continue
        _set_run_fonts(
            style,
            western="Times New Roman",
            east_asia="黑体",
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

    for p in list(doc.element.body.iter(qn("w:p"))):
        if _is_in_table(p):
            continue
        sid = _para_style_id(p)
        if sid in HEADING_STYLE_IDS or sid.startswith("Heading"):
            continue
        if sid in NO_INDENT_STYLE_IDS:
            continue
        if _para_has_display_math(p):
            pPr = _get_or_add(p, "w:pPr")
            jc = _get_or_add(pPr, "w:jc")
            jc.set(qn("w:val"), "center")
            ind = _get_or_add(pPr, "w:ind")
            ind.set(qn("w:firstLineChars"), "0")
            ind.set(qn("w:firstLine"), "0")
            continue
        if _para_is_centered(p):
            continue
        if sid in BODY_STYLE_IDS or sid == "":
            pPr = _get_or_add(p, "w:pPr")
            ind = _get_or_add(pPr, "w:ind")
            ind.set(qn("w:firstLineChars"), "200")
            ind.set(qn("w:firstLine"), "480")
            jc = _get_or_add(pPr, "w:jc")
            jc.set(qn("w:val"), "both")
            spacing = _get_or_add(pPr, "w:spacing")
            spacing.set(qn("w:line"), "360")
            spacing.set(qn("w:lineRule"), "auto")
            spacing.set(qn("w:before"), "0")
            spacing.set(qn("w:after"), "0")

    doc.save(str(path))
