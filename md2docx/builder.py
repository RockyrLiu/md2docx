"""Document builder — assembles a python-docx Document from parsed AST + config.

Covers: cover page, table of contents, headings, paragraphs (with inline
runs), academic three-line tables, images, ordered/unordered lists,
code blocks, blockquotes, thematic breaks, and LaTeX math equations.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml.shared import OxmlElement
from docx.shared import Cm, Pt

from md2docx.config import (
    AppConfig,
    CodeStyle,
    HeadingStyle,
    StyleConfig,
    TableStyle,
    TextStyle,
)
from md2docx.math_handler import add_inline_math, add_math_paragraph
from md2docx.styles import (
    add_cover_line,
    apply_font_to_run,
    apply_format_to_para,
    setup_heading_styles,
    setup_normal_style,
    format_table_content,
    insert_toc_field,
    make_three_line_table,
    set_para_left_border,
    set_para_shading,
    set_run_font,
    SIZE_SANHAO,
)


# =============================================================================
# Inline rendering
# =============================================================================


def _render_inline_run(
    para,
    run_data: dict[str, Any],
    styles: StyleConfig,
    md_path: Path = Path("."),
) -> None:
    """Render a single inline run dict into a paragraph.

    Recursively handles bold, italic, code, links, inline math, images, and
    plain text.
    """
    rtype = run_data.get("type", "text")

    if rtype == "text":
        text = run_data.get("text", "")
        if text:
            run = para.add_run(text)
            apply_font_to_run(run, styles.body)

    elif rtype == "bold":
        children = run_data.get("children", [])
        for child in children:
            run = para.add_run(_extract_run_text(child))
            apply_font_to_run(run, styles.body)
            run.bold = True

    elif rtype == "italic":
        children = run_data.get("children", [])
        for child in children:
            run = para.add_run(_extract_run_text(child))
            apply_font_to_run(run, styles.body)
            run.italic = True

    elif rtype == "code":
        text = run_data.get("text", "")
        if text:
            run = para.add_run(text)
            apply_font_to_run(run, styles.code)

    elif rtype == "link":
        url = run_data.get("url", "")
        children = run_data.get("children", [])
        text = run_data.get("text", _extract_children_text(children))
        if text:
            run = para.add_run(text)
            apply_font_to_run(run, styles.body)
            run.underline = True
            # Store hyperlink via XML
            _add_hyperlink(run, url)

    elif rtype == "inline_math":
        latex = run_data.get("text", "")
        if latex:
            # Try OMML first
            success = add_inline_math(para, latex)
            if not success:
                run = para.add_run(f" ${latex}$ ")
                run.font.name = "Cambria Math"
                run.italic = True

    elif rtype == "image":
        src = run_data.get("src", "")
        alt = run_data.get("alt", "")
        _render_inline_image(para, src, alt, md_path)

    elif rtype == "linebreak":
        run = para.add_run("\n")
    elif rtype == "softbreak":
        run = para.add_run(" ")


def _extract_run_text(run_data: dict[str, Any]) -> str:
    """Get plain text from a run dict (recursive)."""
    rtype = run_data.get("type", "text")
    if rtype == "text":
        return run_data.get("text", "")
    elif rtype in ("bold", "italic", "link"):
        return _extract_children_text(run_data.get("children", []))
    elif rtype == "code":
        return run_data.get("text", "")
    else:
        return run_data.get("text", "")


def _extract_children_text(children: list[dict[str, Any]]) -> str:
    """Extract plain text from children list."""
    parts = []
    for c in children:
        parts.append(_extract_run_text(c))
    return "".join(parts)


def _add_hyperlink(run, url: str) -> None:
    """Wrap a run in a hyperlink element."""
    r_elem = run._element
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), url)  # Simplified — needs relationships for real hyperlinks
    # This is a simplified version; full hyperlinks need relationship management
    # For now we just keep the underlined text


# =============================================================================
# Block renderers
# =============================================================================


def _render_paragraph(doc: Document, block: dict[str, Any], styles: StyleConfig, md_path: Path = Path(".")) -> None:
    """Render a markdown paragraph with inline formatting."""
    para = doc.add_paragraph()
    apply_format_to_para(para, styles.body)

    children = block.get("children", [])
    for run_data in children:
        _render_inline_run(para, run_data, styles, md_path)

    # Clean up empty paragraphs
    if not para.text.strip() and not _para_has_content(para):
        return


def _para_has_content(para) -> bool:
    """Check if paragraph has non-whitespace content or non-text elements."""
    if para.text.strip():
        return True
    # Check for math OMML elements
    for child in para._element:
        if child.tag.endswith("}oMath") or child.tag.endswith("}oMathPara"):
            return True
    return False


def _render_heading(doc: Document, block: dict[str, Any], heading_styles: dict[str, HeadingStyle]) -> None:
    """Render a heading block."""
    level = block.get("level", 1)
    text = block.get("text", "")
    level_key = f"h{max(1, min(6, level))}"
    hs = heading_styles.get(level_key, heading_styles["h1"])
    style_name = f"Heading {max(1, min(6, level))}"

    para = doc.add_paragraph(text, style=style_name)
    for run in para.runs:
        set_run_font(run, cn_name=hs.font_name, en_name=hs.font_name_ascii,
                     size=Pt(hs.font_size), bold=hs.bold)


def _render_table(doc: Document, block: dict[str, Any], table_style: TableStyle) -> None:
    """Render a markdown table as an academic three-line table."""
    header = block.get("header", [])
    rows = block.get("rows", [])

    if not header and not rows:
        return

    # Determine column count
    col_count = max(len(header), max((len(r) for r in rows), default=0))
    if col_count == 0:
        return

    # Normalize
    header = header + [""] * (col_count - len(header))
    norm_rows = [r + [""] * (col_count - len(r)) for r in rows]

    table = doc.add_table(rows=1 + len(norm_rows), cols=col_count)
    table.autofit = True

    # Header row
    _fill_table_row(table.rows[0], header, table_style, is_header=True)

    # Data rows
    for i, row_data in enumerate(norm_rows):
        _fill_table_row(table.rows[i + 1], row_data, table_style, is_header=False)

    # Apply three-line formatting
    make_three_line_table(table)
    format_table_content(table, table_style)


def _fill_table_row(row, cell_texts: list[str], table_style: TableStyle, *, is_header: bool) -> None:
    """Fill a table row with text."""
    for i, text in enumerate(cell_texts):
        cell = row.cells[i]
        cell.text = ""
        para = cell.paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.line_spacing = table_style.line_spacing
        run = para.add_run(text.strip())
        apply_font_to_run(run, table_style)
        if is_header and table_style.header_bold:
            run.bold = True


def _render_image(doc: Document, block: dict[str, Any], md_path: Path) -> None:
    """Render an image block."""
    src = block.get("src", "")
    alt = block.get("alt", "")

    # Resolve path relative to markdown file
    img_path = md_path.parent / src if not Path(src).is_absolute() else Path(src)

    if not img_path.exists():
        # Add placeholder text
        para = doc.add_paragraph()
        run = para.add_run(f"[图片缺失: {src}]")
        run.font.color.rgb = None  # default
        run.italic = True
        return

    from docx.shared import Inches as DocxInches

    try:
        from PIL import Image

        with Image.open(img_path) as img:
            # Calculate size to fit within page (max 5.5 inches for A4 with 3.18cm margins)
            max_width = DocxInches(5.5)
            max_height = DocxInches(7.0)
            aspect = img.width / img.height
            width = max_width
            height = DocxInches(width.inches / aspect)
            if height > max_height:
                height = max_height
                width = DocxInches(height.inches * aspect)

        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run()
        run.add_picture(str(img_path), width=width, height=height)

        # Alt text caption
        if alt:
            cap_para = doc.add_paragraph()
            cap_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cap_run = cap_para.add_run(f"图：{alt}")
            cap_run.font.size = Pt(10)
            cap_run.italic = True

    except Exception:
        para = doc.add_paragraph()
        run = para.add_run(f"[无法加载图片: {src}]")
        run.italic = True


def _render_inline_image(para, src: str, alt: str, md_path: Path) -> None:
    """Render an image inline within a paragraph."""
    img_path = md_path.parent / src if not Path(src).is_absolute() else Path(src)

    if not img_path.exists():
        run = para.add_run(f"[图片缺失: {src}]")
        run.font.italic = True
        return

    from docx.shared import Inches as DocxInches

    try:
        from PIL import Image

        with Image.open(img_path) as img:
            max_width = DocxInches(2.0)  # Inline images smaller than block
            aspect = img.width / img.height
            width = max_width
            height = DocxInches(width.inches / aspect)
            if height > DocxInches(2.0):
                height = DocxInches(2.0)
                width = DocxInches(height.inches * aspect)

        run = para.add_run()
        if alt:
            before = para.add_run(f"[{alt}] ")
            before.font.size = Pt(9)
            before.font.italic = True
        run.add_picture(str(img_path), width=width, height=height)
    except Exception:
        run = para.add_run(f"[无法加载图片: {src}]")
        run.font.italic = True


def _render_list(doc: Document, block: dict[str, Any], styles: StyleConfig, md_path: Path = Path("."), level: int = 0) -> None:
    """Render an ordered or unordered list."""
    list_style = styles.list
    ordered = block.get("ordered", False)
    items = block.get("items", [])

    # Build a StyleConfig where body uses list-style fonts (for _render_inline_run)
    list_body = TextStyle(
        font_name=list_style.font_name,
        font_name_ascii=list_style.font_name_ascii,
        font_size=list_style.font_size,
        line_spacing=list_style.line_spacing,
    )
    list_styles = StyleConfig(body=list_body)

    for idx, item in enumerate(items):
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        para.paragraph_format.line_spacing = list_style.line_spacing
        para.paragraph_format.left_indent = Cm(list_style.indent * 2.54 * (level + 1))

        # Bullet or number prefix
        if ordered:
            prefix = f"{idx + 1}. "
        else:
            prefix = "• "

        prefix_run = para.add_run(prefix)
        apply_font_to_run(prefix_run, list_style)

        # Render inline runs (preserves math, bold, italic, etc.)
        inline_runs = item.get("inline_runs", [])
        if inline_runs:
            for run_data in inline_runs:
                _render_inline_run(para, run_data, list_styles, md_path)
        else:
            # Fallback for plain text (backward compat)
            text = item.get("text", "")
            if text:
                run = para.add_run(text)
                apply_font_to_run(run, list_style)

        # Handle nested lists
        for child in item.get("children", []):
            if isinstance(child, dict) and child.get("type") == "list":
                _render_list(doc, child, styles, md_path, level + 1)


def _render_code_block(doc: Document, block: dict[str, Any], code_style: CodeStyle) -> None:
    """Render a fenced code block."""
    text = block.get("text", "")
    para = doc.add_paragraph()

    set_para_shading(para, code_style.background_color)
    para.paragraph_format.space_before = Pt(3)
    para.paragraph_format.space_after = Pt(3)
    para.paragraph_format.left_indent = Cm(0.5)
    para.paragraph_format.line_spacing = code_style.line_spacing

    for line in text.split("\n"):
        if para.text:
            # Add line break for multi-line
            run = para.add_run("\n" + line)
        else:
            run = para.add_run(line)
        apply_font_to_run(run, code_style)


def _render_blockquote(doc: Document, block: dict[str, Any], styles: StyleConfig, md_path: Path = Path(".")) -> None:
    """Render a blockquote."""
    children = block.get("children", [])
    for child in children:
        ctype = child.get("type", "")
        if ctype == "paragraph":
            para = doc.add_paragraph()
            set_para_left_border(para, styles.blockquote.border_left_color, 12)
            para.paragraph_format.left_indent = Cm(styles.blockquote.left_indent * 2.54)
            para.paragraph_format.line_spacing = styles.blockquote.line_spacing
            for run_data in child.get("children", []):
                _render_inline_run(para, run_data, styles, md_path)
            # Style runs
            for run in para.runs:
                apply_font_to_run(run, styles.blockquote)
        elif ctype == "heading":
            level = child.get("level", 1)
            hs = styles.headings.get(f"h{level}", styles.headings["h1"])
            para = doc.add_paragraph(child.get("text", ""))
            set_para_left_border(para, styles.blockquote.border_left_color, 12)
            para.paragraph_format.left_indent = Cm(styles.blockquote.left_indent * 2.54)
            for run in para.runs:
                apply_font_to_run(run, hs)
        else:
            _render_block(doc, child, styles, md_path)


# =============================================================================
# Page setup
# =============================================================================


def _setup_page(doc: Document, config: AppConfig) -> None:
    """Configure page size and margins."""
    for section in doc.sections:
        # Page size
        if config.page.size == "Letter":
            section.page_width = Cm(21.59)
            section.page_height = Cm(27.94)
        else:
            section.page_width = Cm(21.0)
            section.page_height = Cm(29.7)

        # Margins
        section.top_margin = Cm(config.page.top_margin)
        section.bottom_margin = Cm(config.page.bottom_margin)
        section.left_margin = Cm(config.page.left_margin)
        section.right_margin = Cm(config.page.right_margin)


# =============================================================================
# Cover page
# =============================================================================


def _build_cover(doc: Document, config: AppConfig) -> None:
    cover = config.cover
    # Vertical spacing from top
    for _ in range(4):
        doc.add_paragraph()

    # Report Title
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_before = Pt(20)
    run = title_p.add_run(cover.title)
    run.font.name = '黑体'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
    run.font.size = Pt(28)
    run.bold = True

    # Spacing
    for _ in range(4):
        doc.add_paragraph()

    # Student info
    add_cover_line(doc, '姓    名：', f'{cover.author}')
    add_cover_line(doc, '班    级：', f'{cover.class_info}')
    add_cover_line(doc, '学    号：', f'{cover.student_id}')
    add_cover_line(doc, '任课教师：', f'{cover.teacher}')
    add_cover_line(doc, '实验日期：', f'{cover.date}')

    doc.add_page_break()

# =============================================================================
# TOC
# =============================================================================


def _build_toc(doc: Document, config: AppConfig) -> None:
    """Insert a Word Table of Contents field."""
    # "目录" heading
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.paragraph_format.space_before = Pt(12)
    para.paragraph_format.space_after = Pt(12)
    run = para.add_run("目  录")
    set_run_font(run, cn_name="黑体", size=SIZE_SANHAO, bold=True)

    # TOC field
    insert_toc_field(doc, config.toc.level)

    # Page break after TOC
    doc.add_page_break()


# =============================================================================
# Block dispatcher
# =============================================================================


def _render_block(doc: Document, block: dict[str, Any], styles: StyleConfig, md_path: Path) -> None:
    """Dispatch a single block to the appropriate renderer."""
    btype = block.get("type", "")

    if btype == "heading":
        _render_heading(doc, block, styles.headings)

    elif btype == "paragraph":
        _render_paragraph(doc, block, styles, md_path)

    elif btype == "table":
        _render_table(doc, block, styles.table)

    elif btype == "image":
        _render_image(doc, block, md_path)

    elif btype == "list":
        _render_list(doc, block, styles, md_path)

    elif btype == "code_block":
        _render_code_block(doc, block, styles.code)

    elif btype == "blockquote":
        _render_blockquote(doc, block, styles, md_path)

    elif btype == "thematic_break":
        doc.add_paragraph("—" * 30)

    elif btype == "blank_line":
        pass  # skip — paragraph spacing handles separation

    elif btype == "block_math":
        latex = block.get("text", "")
        add_math_paragraph(doc, latex)


# =============================================================================
# Main entry point
# =============================================================================


def build_docx(
    ast: list[dict[str, Any]],
    config: AppConfig,
    md_path: str | Path,
) -> Document:
    """Build a python-docx Document from a parsed AST and configuration.

    Parameters
    ----------
    ast : list[dict]
        Parsed markdown block list from ``parse_markdown()``.
    config : AppConfig
        Full application configuration.
    md_path : str | Path
        Path to the source markdown file (for resolving image paths).

    Returns
    -------
    Document
        The built python-docx Document object (before saving).
    """
    md_path = Path(md_path)

    doc = Document()

    # --- Page setup ---
    _setup_page(doc, config)

    # --- Normal style ---
    setup_normal_style(doc, config.styles.body)

    # --- Heading styles ---
    setup_heading_styles(doc, config.styles.headings)

    # --- Cover page ---
    if config.cover.enabled:
        _build_cover(doc, config)

    # --- Table of Contents ---
    if config.toc.enabled:
        _build_toc(doc, config)

    # --- Content ---
    for block in ast:
        _render_block(doc, block, config.styles, md_path)

    return doc
