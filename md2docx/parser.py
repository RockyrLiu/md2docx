"""Markdown parser — converts .md files to a flat block-level AST using mistune v3.

Handles: headings, paragraphs, tables, images, ordered/unordered lists
(both tight and loose, nested), blockquotes, fenced code blocks,
thematic breaks, inline formatting (bold, italic, code, links), and
LaTeX-style math ($inline$ and $$display$$).
"""

from __future__ import annotations

import re
from typing import Any

import mistune

# ---------------------------------------------------------------------------
# Inline / block math regexes
# ---------------------------------------------------------------------------
_INLINE_MATH_RE = re.compile(r"(?<!\\)\$((?:[^\$]|\\\$)+?)(?<!\\)\$")
_DISPLAY_MATH_RE = re.compile(r"(?<!\\)\$\$((?:[^\$]|\\\$)+?)(?<!\\)\$\$", re.DOTALL)
_PH_MATH_INLINE = re.compile(r"\x00MATHINLINE\d+\x00")
_PH_MATH_BLOCK = re.compile(r"\x00MATHBLOCK\d+\x00")
_PH_ANY = re.compile(r"\x00MATH(?:INLINE|BLOCK)\d+\x00")


def _escape_math(text: str) -> tuple[str, dict[str, str]]:
    """Temporarily replace math spans with placeholders so mistune won't
    mangle them.  Returns (processed_text, placeholder_map)."""
    placeholders: dict[str, str] = {}
    counter = 0

    def _replace_display(m: re.Match) -> str:
        nonlocal counter
        key = f"\x00MATHBLOCK{counter}\x00"
        placeholders[key] = m.group(1)  # content without $$ signs
        counter += 1
        return f"\n\n{key}\n\n"

    def _replace_inline(m: re.Match) -> str:
        nonlocal counter
        key = f"\x00MATHINLINE{counter}\x00"
        placeholders[key] = m.group(1)
        counter += 1
        return f" {key} "

    text = _DISPLAY_MATH_RE.sub(_replace_display, text)
    text = _INLINE_MATH_RE.sub(_replace_inline, text)
    return text, placeholders


def _restore_text(text: str, placeholders: dict[str, str]) -> str:
    """Restore math placeholders in a text string, returning the raw math."""
    for key, value in placeholders.items():
        text = text.replace(key, value)
    return text


# ---------------------------------------------------------------------------
# Inline children extraction
# ---------------------------------------------------------------------------


def _split_by_placeholders(text: str, placeholders: dict[str, str]) -> list[dict[str, Any]]:
    """Split text on math placeholder markers, producing text + inline_math runs."""
    runs: list[dict[str, Any]] = []
    parts = _PH_ANY.split(text)
    markers = _PH_ANY.findall(text)

    # parts and markers interleave: part0, marker0, part1, marker1, ...
    for i, part in enumerate(parts):
        if part.strip():
            runs.append({"type": "text", "text": part})
        if i < len(markers):
            marker = markers[i]
            math_content = placeholders.get(marker, marker)
            runs.append({"type": "inline_math", "text": math_content})
    return runs


def _extract_inline_children(
    children: list[dict[str, Any]],
    placeholders: dict[str, str],
) -> list[dict[str, Any]]:
    """Convert mistune v3 inline AST children into our run-dict format.

    mistune v3 uses ``"raw"`` for text nodes (not ``"text"``).
    """
    result: list[dict[str, Any]] = []

    for child in children:
        if not isinstance(child, dict):
            continue
        ctype = child.get("type", "")

        if ctype == "text":
            raw = child.get("raw", "")
            if not raw:
                continue
            # Check for math placeholders BEFORE restoring
            if _PH_ANY.search(raw):
                result.extend(_split_by_placeholders(raw, placeholders))
            else:
                result.append({"type": "text", "text": raw})

        elif ctype == "strong":
            inner = _extract_inline_children(child.get("children", []), placeholders)
            result.append({"type": "bold", "children": inner})

        elif ctype == "emphasis":
            inner = _extract_inline_children(child.get("children", []), placeholders)
            result.append({"type": "italic", "children": inner})

        elif ctype == "codespan":
            raw = child.get("raw", "")
            result.append({"type": "code", "text": raw})

        elif ctype == "link":
            inner = _extract_inline_children(child.get("children", []), placeholders)
            result.append({
                "type": "link",
                "url": child.get("link", ""),
                "children": inner,
            })

        elif ctype == "image":
            # mistune v3: attrs.url for src, children[text] for alt
            src = child.get("attrs", {}).get("url", "")
            alt = _extract_plain_text(child.get("children", []))
            result.append({"type": "image", "src": src, "alt": alt})

        elif ctype == "linebreak":
            result.append({"type": "linebreak"})
        elif ctype == "softbreak":
            result.append({"type": "softbreak"})
        elif ctype == "inline_html":
            raw = child.get("raw", "")
            result.append({"type": "text", "text": raw})

    return result


def _split_math_runs(text: str, placeholders: dict[str, str]) -> list[dict[str, Any]]:
    """Split text containing $...$ spans into text + inline_math runs."""
    runs: list[dict[str, Any]] = []
    pos = 0
    for m in _INLINE_MATH_RE.finditer(text):
        if m.start() > pos:
            prefix = text[pos : m.start()]
            if prefix.strip():
                runs.append({"type": "text", "text": prefix})
        runs.append({"type": "inline_math", "text": m.group(1)})
        pos = m.end()
    if pos < len(text):
        suffix = text[pos:]
        if suffix.strip():
            runs.append({"type": "text", "text": suffix})
    return runs


# ---------------------------------------------------------------------------
# Plain text extraction (for headings / table cells / list items)
# ---------------------------------------------------------------------------


def _extract_plain_text(children: list[dict[str, Any]]) -> str:
    """Recursively extract plain text from mistune inline children."""
    parts: list[str] = []
    for child in children:
        if not isinstance(child, dict):
            continue
        ctype = child.get("type", "")
        if ctype == "text":
            parts.append(child.get("raw", ""))
        elif ctype in ("strong", "emphasis", "link"):
            parts.append(_extract_plain_text(child.get("children", [])))
        elif ctype == "codespan":
            parts.append(child.get("raw", ""))
        elif ctype in ("linebreak", "softbreak"):
            parts.append(" ")
        elif ctype == "image":
            parts.append(child.get("alt", ""))
    return "".join(parts)


def _extract_plain_text_from_block_text(children: list[dict[str, Any]]) -> str:
    """Extract text from block_text children (used in tight lists)."""
    parts: list[str] = []
    for child in children:
        if not isinstance(child, dict):
            continue
        ctype = child.get("type", "")
        if ctype == "block_text":
            parts.append(_extract_plain_text(child.get("children", [])))
        elif ctype == "paragraph":
            parts.append(_extract_plain_text(child.get("children", [])))
        elif ctype == "list":
            # nested list — skip for text extraction
            pass
        elif ctype == "text":
            parts.append(child.get("raw", ""))
    return "".join(parts)


# ---------------------------------------------------------------------------
# Block conversion
# ---------------------------------------------------------------------------


def _convert_table_cell(cell: dict[str, Any]) -> list[dict[str, Any]]:
    """Extract inline children from a table_cell."""
    return cell.get("children", [])


def _block_to_dict(
    block: dict[str, Any],
    placeholders: dict[str, str],
) -> dict[str, Any] | None:
    """Convert a single mistune v3 AST block to our intermediate format."""
    btype = block.get("type", "")

    # --- Thematic break ---
    if btype == "thematic_break":
        return {"type": "thematic_break"}

    # --- Heading ---
    if btype == "heading":
        level = block.get("attrs", {}).get("level", 1)
        text = _extract_plain_text(block.get("children", []))
        text = _restore_text(text, placeholders)
        return {"type": "heading", "level": int(level), "text": text.strip()}

    # --- Paragraph ---
    if btype == "paragraph":
        children = block.get("children", [])
        runs = _extract_inline_children(children, placeholders)
        # Keep paragraphs even if empty (they may be spacers)
        return {"type": "paragraph", "children": runs}

    # --- Blank line ---
    if btype == "blank_line":
        return {"type": "blank_line"}

    # --- Fenced code block ---
    if btype == "block_code":
        return {
            "type": "code_block",
            "text": block.get("raw", ""),
            "language": block.get("attrs", {}).get("info", ""),
        }

    # --- Block quote ---
    if btype == "block_quote":
        inner_blocks = _flatten_blocks(block.get("children", []), placeholders)
        return {"type": "blockquote", "children": inner_blocks}

    # --- Table (mistune v3: table → table_head/table_body → table_row → table_cell) ---
    if btype == "table":
        header_cells: list[str] = []
        body_rows: list[list[str]] = []

        for section in block.get("children", []):
            stype = section.get("type", "")
            if stype == "table_head":
                # table_head children are table_cell directly (no table_row)
                for cell in section.get("children", []):
                    text = _extract_plain_text(cell.get("children", []))
                    text = _restore_text(text, placeholders)
                    header_cells.append(text)
            elif stype == "table_body":
                # table_body children are table_row → table_cell
                for row in section.get("children", []):
                    row_texts: list[str] = []
                    for cell in row.get("children", []):
                        text = _extract_plain_text(cell.get("children", []))
                        text = _restore_text(text, placeholders)
                        row_texts.append(text)
                    body_rows.append(row_texts)

        return {
            "type": "table",
            "header": header_cells,
            "rows": body_rows,
        }

    # --- List ---
    if btype == "list":
        ordered = block.get("attrs", {}).get("ordered", False)
        items: list[dict[str, Any]] = []

        for item in block.get("children", []):
            item_type = item.get("type", "")
            if item_type != "list_item":
                continue

            item_children = item.get("children", [])
            # Extract inline runs (preserves bold, italic, math, etc.) and nested lists
            inline_runs: list[dict[str, Any]] = []
            sub_lists: list[dict[str, Any]] = []

            for ic in item_children:
                ic_type = ic.get("type", "")
                if ic_type == "block_text":
                    runs = _extract_inline_children(ic.get("children", []), placeholders)
                    inline_runs.extend(runs)
                elif ic_type == "list":
                    sub = _block_to_dict(ic, placeholders)
                    if sub:
                        sub_lists.append(sub)
                elif ic_type == "paragraph":
                    runs = _extract_inline_children(ic.get("children", []), placeholders)
                    inline_runs.extend(runs)

            items.append({
                "type": "list_item",
                "inline_runs": inline_runs,
                "children": sub_lists,
            })

        return {"type": "list", "ordered": ordered, "items": items}

    # --- Block HTML (may contain <img>) ---
    if btype == "block_html":
        raw = block.get("raw", "")
        img_match = re.search(r'<img[^>]+src=["\']([^"\']+)["\']', raw)
        if img_match:
            return {"type": "image", "src": img_match.group(1), "alt": ""}
        return None

    return None


def _flatten_blocks(
    blocks: list[dict[str, Any]],
    placeholders: dict[str, str],
) -> list[dict[str, Any]]:
    """Convert a list of mistune v3 AST blocks into our flat format."""
    result: list[dict[str, Any]] = []
    for block in blocks:
        if not isinstance(block, dict):
            continue

        # Check if this is a display math placeholder (standalone paragraph)
        btype = block.get("type", "")
        if btype == "paragraph":
            raw_text = _extract_plain_text(block.get("children", []))
            for key, value in placeholders.items():
                if key in raw_text and key.startswith("\x00MATHBLOCK"):
                    result.append({"type": "block_math", "text": value})
                    # Don't add as paragraph
                    break
            else:
                converted = _block_to_dict(block, placeholders)
                if converted is None:
                    pass
                elif isinstance(converted, dict) and converted.get("type") == "paragraph":
                    # Standalone image(s) in a paragraph → image blocks
                    para_children = converted.get("children", [])
                    if para_children and all(c.get("type") == "image" for c in para_children):
                        result.extend(para_children)
                    else:
                        result.append(converted)
                else:
                    result.append(converted)
        else:
            converted = _block_to_dict(block, placeholders)
            if converted is not None:
                result.append(converted)

    return result


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def parse_markdown(text: str) -> list[dict[str, Any]]:
    """Parse Markdown text into a flat list of block-level element dicts.

    Parameters
    ----------
    text : str
        Raw Markdown source.

    Returns
    -------
    list[dict]
        Each dict has at least ``{"type": str}``.
    """
    # 1. Escape math sections → placeholders
    escaped, placeholders = _escape_math(text)

    # 2. Parse with mistune (renderer='ast' for v3)
    markdown = mistune.create_markdown(
        renderer="ast",
        plugins=["table", "task_lists"],
    )
    ast = markdown(escaped)

    # 3. Flatten to our format, restoring math
    return _flatten_blocks(ast, placeholders)
