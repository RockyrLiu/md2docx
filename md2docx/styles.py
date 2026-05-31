"""Word style helpers — clean, readable API for docx formatting.

Inspired by the direct, self-documenting style of standalone report scripts.
All font sizes follow the Chinese academic convention (中文字号).
"""

from __future__ import annotations

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml.shared import OxmlElement
from docx.shared import Cm, Pt, RGBColor

# =============================================================================
# Chinese font size constants (中文字号 → pt)
# =============================================================================
SIZE_CHUHAO = Pt(42)       # 初号
SIZE_XIAOCHU = Pt(36)      # 小初
SIZE_YIHAO = Pt(26)        # 一号
SIZE_XIAOYI = Pt(24)       # 小一
SIZE_ERHAO = Pt(22)        # 二号
SIZE_XIAOER = Pt(18)       # 小二
SIZE_SANHAO = Pt(16)       # 三号
SIZE_XIAOSAN = Pt(15)      # 小三
SIZE_SIHAO = Pt(14)        # 四号
SIZE_XIAOSI = Pt(12)       # 小四（正文）
SIZE_WUHAO = Pt(10.5)      # 五号（表格/脚注）
SIZE_XIAOWU = Pt(9)        # 小五

# =============================================================================
# Alignment map
# =============================================================================
_ALIGN = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

# =============================================================================
# Run (character) formatting
# =============================================================================


def set_run_font(
    run,
    cn_name: str = "宋体",
    en_name: str = "Times New Roman",
    size: Pt = SIZE_XIAOSI,
    bold: bool = False,
    italic: bool = False,
    color: RGBColor | None = None,
) -> None:
    """Apply font properties to a run — one call, self-documenting.

    Usage::

        set_run_font(run, cn_name="黑体", size=SIZE_SANHAO, bold=True)
    """
    # Set ASCII/HAnsi via python-docx API (creates rPr + rFonts if needed)
    run.font.name = en_name
    run.font.size = size
    run.bold = bold
    run.italic = italic

    # Clear theme references and set East-Asian font explicitly
    rpr = run._element.rPr
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    _clear_theme_fonts(rfonts)
    rfonts.set(qn("w:eastAsia"), cn_name)

    # Color: always explicit (no theme)
    if color is not None:
        _clear_theme_color(rpr)
        run.font.color.rgb = color


# =============================================================================
# Paragraph formatting
# =============================================================================


def set_para_format(
    para,
    line_spacing: float = 1.5,
    alignment: str = "left",
    first_line_indent_chars: int = 0,
    font_size_for_indent: Pt | None = None,
    space_before_pt: int = 0,
    space_after_pt: int = 0,
) -> None:
    """Apply paragraph-level formatting — one call, self-documenting.

    Usage::

        set_para_format(para, line_spacing=1.5, alignment="justify",
                        first_line_indent_chars=2, font_size_for_indent=SIZE_XIAOSI)
    """
    para.paragraph_format.line_spacing = line_spacing

    if alignment in _ALIGN:
        para.alignment = _ALIGN[alignment]

    if first_line_indent_chars > 0 and font_size_for_indent is not None:
        # 1 Chinese char ≈ font_size in pt → cm
        indent_cm = first_line_indent_chars * font_size_for_indent.pt * 0.0353
        para.paragraph_format.first_line_indent = Cm(indent_cm)

    if space_before_pt:
        para.paragraph_format.space_before = Pt(space_before_pt)
    if space_after_pt:
        para.paragraph_format.space_after = Pt(space_after_pt)


# =============================================================================
# Paragraph shading / borders (code blocks, blockquotes)
# =============================================================================


def set_para_shading(para, hex_color: str) -> None:
    """Set paragraph background color (e.g. for code blocks)."""
    ppr = para._element.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    ppr.append(shd)


def set_para_left_border(para, hex_color: str = "999999", width_eighths_pt: int = 12) -> None:
    """Add a left border line to a paragraph (e.g. for blockquotes)."""
    ppr = para._element.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), str(width_eighths_pt))
    left.set(qn("w:space"), "4")
    left.set(qn("w:color"), hex_color)
    p_bdr.append(left)
    ppr.append(p_bdr)


# =============================================================================
# Apply a style-config object to a run / paragraph
# =============================================================================


def apply_font_to_run(run, style) -> None:
    """Apply font properties from a TextStyle/HeadingStyle/TableStyle config to a run."""
    cn = style.font_name
    en = getattr(style, "font_name_ascii", cn)
    color = None
    if getattr(style, "color", None):
        try:
            color = RGBColor.from_string(style.color)
        except Exception:
            pass
    set_run_font(
        run,
        cn_name=cn,
        en_name=en,
        size=Pt(style.font_size),
        bold=getattr(style, "bold", False),
        italic=getattr(style, "italic", False),
        color=color,
    )


def apply_format_to_para(para, style) -> None:
    """Apply paragraph formatting from a TextStyle config (spacing, indent, alignment)."""
    set_para_format(
        para,
        line_spacing=getattr(style, "line_spacing", 1.5),
        alignment=getattr(style, "alignment", "left"),
        first_line_indent_chars=getattr(style, "first_line_indent", 0),
        font_size_for_indent=Pt(getattr(style, "font_size", 12)),
        space_before_pt=getattr(style, "space_before", 0),
        space_after_pt=getattr(style, "space_after", 0),
    )


# =============================================================================
# Cover page helpers
# =============================================================================


def add_cover_title(doc: Document, text: str) -> None:
    """Add a large, centered title on the cover page (黑体 一号)."""
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.paragraph_format.space_before = Pt(60)
    para.paragraph_format.space_after = Pt(30)
    run = para.add_run(text)
    set_run_font(run, cn_name="黑体", en_name="Times New Roman",
                 size=SIZE_YIHAO, bold=True)
    return para


def add_cover_info_line(
    doc: Document,
    label: str,
    value: str = "",
    size: Pt = SIZE_SANHAO,
) -> None:
    """Add a centered info line on the cover: ``label：value``."""
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.paragraph_format.line_spacing = 1.5
    para.paragraph_format.space_before = Pt(6)
    para.paragraph_format.space_after = Pt(6)
    text = f"{label}：{value}" if value else label
    run = para.add_run(text)
    set_run_font(run, cn_name="宋体", size=size)
    return para


# =============================================================================
# Table of Contents
# =============================================================================


def insert_toc_field(doc: Document, levels: int = 3) -> None:
    """Insert a Word TOC field (requires right-click → Update Field in Word)."""
    para = doc.add_paragraph()

    run = para.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    run._element.append(fld_begin)

    run2 = para.add_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f'TOC \\o "1-{levels}" \\h \\z \\u'
    run2._element.append(instr)

    run3 = para.add_run()
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    run3._element.append(fld_sep)

    hint = para.add_run("（请在 Word 中右键此处 → 更新域，以生成目录）")
    set_run_font(hint, cn_name="宋体", size=SIZE_XIAOSI, italic=True,
                 color=RGBColor(128, 128, 128))

    run4 = para.add_run()
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run4._element.append(fld_end)


# =============================================================================
# Style setup (Normal + Heading 1-6)
# =============================================================================


def _clear_theme_fonts(rfonts) -> None:
    """Remove theme font references so explicit fonts take effect."""
    for attr in (qn("w:asciiTheme"), qn("w:eastAsiaTheme"), qn("w:hAnsiTheme"), qn("w:cstheme")):
        try:
            del rfonts.attrib[attr]
        except KeyError:
            pass


def _clear_theme_color(rpr) -> None:
    """Remove theme color reference from rPr so explicit color takes effect."""
    color_el = rpr.find(qn("w:color"))
    if color_el is not None:
        try:
            del color_el.attrib[qn("w:themeColor")]
        except KeyError:
            pass
        try:
            del color_el.attrib[qn("w:themeShade")]
        except KeyError:
            pass


def _set_style_fonts(word_style, cn_name: str, en_name: str) -> None:
    """Fully override style-level fonts — clear themes, set explicit CJK + ASCII."""
    rpr = word_style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)

    # Remove inherited theme font references (they take precedence over explicit names)
    _clear_theme_fonts(rfonts)

    # Set all font slots explicitly
    rfonts.set(qn("w:ascii"), en_name)
    rfonts.set(qn("w:hAnsi"), en_name)
    rfonts.set(qn("w:eastAsia"), cn_name)
    rfonts.set(qn("w:cs"), en_name)

    # Also set via the python-docx Font API (duplicates ascii/hAnsi but keeps consistency)
    word_style.font.name = en_name


def _set_style_color(word_style, hex_color: str | None) -> None:
    """Set explicit font color on a style, clearing any theme color."""
    rpr = word_style.element.get_or_add_rPr()
    _clear_theme_color(rpr)

    if hex_color:
        try:
            word_style.font.color.rgb = RGBColor.from_string(hex_color)
        except Exception:
            word_style.font.color.rgb = RGBColor(0, 0, 0)
    else:
        word_style.font.color.rgb = RGBColor(0, 0, 0)


def configure_normal_style(doc: Document, body_style) -> None:
    """Override the Normal style with configured body font (宋体 小四)."""
    word_style = doc.styles["Normal"]
    cn = getattr(body_style, "font_name", "宋体")
    en = getattr(body_style, "font_name_ascii", "Times New Roman")

    _set_style_fonts(word_style, cn, en)
    word_style.font.size = Pt(getattr(body_style, "font_size", 12))
    word_style.font.bold = getattr(body_style, "bold", False)
    word_style.font.italic = getattr(body_style, "italic", False)
    _set_style_color(word_style, getattr(body_style, "color", None))

    pf = word_style.paragraph_format
    pf.line_spacing = getattr(body_style, "line_spacing", 1.5)
    align = getattr(body_style, "alignment", "justify")
    if align in _ALIGN:
        pf.alignment = _ALIGN[align]


def configure_heading_styles(doc: Document, heading_config: dict) -> None:
    """Override Word heading styles 1–6 — clear themes, set 黑体 + explicit black.

    heading_config maps 'h1'..'h6' to HeadingStyle objects.
    """
    for i, key in enumerate(("h1", "h2", "h3", "h4", "h5", "h6"), 1):
        hs = heading_config.get(key)
        if hs is None:
            continue

        style_name = f"Heading {i}"

        # Ensure the style exists
        try:
            doc.styles[style_name]
        except KeyError:
            s = doc.styles.add_style(style_name, 1)
            s.base_style = doc.styles["Normal"]

        word_style = doc.styles[style_name]

        # Font: clear themes, set explicit CJK + ASCII
        cn = getattr(hs, "font_name", "黑体")
        en = getattr(hs, "font_name_ascii", "Times New Roman")
        _set_style_fonts(word_style, cn, en)

        # Size / bold / italic
        word_style.font.size = Pt(getattr(hs, "font_size", 14))
        word_style.font.bold = getattr(hs, "bold", True)
        word_style.font.italic = getattr(hs, "italic", False)

        # Color: always explicit (black by default, not theme blue)
        _set_style_color(word_style, getattr(hs, "color", None))

        # Paragraph format
        pf = word_style.paragraph_format
        align = getattr(hs, "alignment", "left")
        if align in _ALIGN:
            pf.alignment = _ALIGN[align]
        pf.space_before = Pt(getattr(hs, "space_before", 6))
        pf.space_after = Pt(getattr(hs, "space_after", 6))


# =============================================================================
# Academic three-line table
# =============================================================================


def make_three_line_table(table, header_rows: int = 1) -> None:
    """Format a python-docx Table as an academic three-line table.

    - Top border: 1.5 pt
    - Below header: 0.75 pt
    - Bottom border: 1.5 pt
    - No vertical or interior horizontal borders
    """
    _clear_all_cell_borders(table)
    _set_table_border(table, "top", 12)       # 1.5 pt
    _set_table_border(table, "bottom", 12)    # 1.5 pt
    _set_table_border(table, "left", 0, "FFFFFF")
    _set_table_border(table, "right", 0, "FFFFFF")
    _set_table_border(table, "insideH", 0, "FFFFFF")
    _set_table_border(table, "insideV", 0, "FFFFFF")

    for i in range(header_rows):
        for cell in table.rows[i].cells:
            _set_cell_border(cell, "bottom", 6)  # 0.75 pt


def format_table_content(table, table_style, header_rows: int = 1) -> None:
    """Apply font, size, alignment to all cells in a table."""
    for row_idx, row in enumerate(table.rows):
        for cell in row.cells:
            for para in cell.paragraphs:
                para.alignment = _ALIGN.get(
                    getattr(table_style, "alignment", "center"),
                    WD_ALIGN_PARAGRAPH.CENTER,
                )
                para.paragraph_format.line_spacing = getattr(table_style, "line_spacing", 1.0)
                for run in para.runs:
                    apply_font_to_run(run, table_style)
                    if row_idx < header_rows and getattr(table_style, "header_bold", True):
                        run.bold = True


# =============================================================================
# Internal: XML-level table border helpers
# =============================================================================


def _clear_all_cell_borders(table) -> None:
    for row in table.rows:
        for cell in row.cells:
            for pos in ("top", "bottom", "left", "right"):
                _set_cell_border(cell, pos, 0, "FFFFFF")


def _set_cell_border(cell, position: str, sz: int, color: str = "000000") -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)

    elem = OxmlElement(f"w:{position}")
    elem.set(qn("w:val"), "single")
    elem.set(qn("w:sz"), str(sz))
    elem.set(qn("w:space"), "0")
    elem.set(qn("w:color"), color)

    existing = borders.find(qn(f"w:{position}"))
    if existing is not None:
        borders.remove(existing)
    borders.append(elem)


def _set_table_border(table, position: str, sz: int, color: str = "000000") -> None:
    tbl_pr = table._tbl.tblPr
    if tbl_pr is None:
        tbl_pr = OxmlElement("w:tblPr")
        table._tbl.insert(0, tbl_pr)

    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.insert(0, borders)

    elem = OxmlElement(f"w:{position}")
    elem.set(qn("w:val"), "single")
    elem.set(qn("w:sz"), str(sz))
    elem.set(qn("w:space"), "0")
    elem.set(qn("w:color"), color)

    existing = borders.find(qn(f"w:{position}"))
    if existing is not None:
        borders.remove(existing)
    borders.append(elem)


