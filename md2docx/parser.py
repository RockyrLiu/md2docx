"""Markdown parser — converts .md files to a flat block-level AST using mistune.

Handles: headings, paragraphs, tables, images, ordered/unordered lists
(both tight and loose, nested), blockquotes, fenced code blocks,
thematic breaks, inline formatting (bold, italic, code, links), and
LaTeX-style math ($inline$ and $$display$$).
"""

from __future__ import annotations

import re
from typing import Any
from urllib.parse import unquote

import mistune

# ---------------------------------------------------------------------------
# Inline / block math regexes
# ---------------------------------------------------------------------------
_INLINE_MATH_RE = re.compile(r"(?<!\\)\$((?:[^\$]|\\\$)+?)(?<!\\)\$")
_DISPLAY_MATH_RE = re.compile(r"(?<!\\)\$\$((?:[^\$]|\\\$)+?)(?<!\\)\$\$", re.DOTALL)
_PH_ANY = re.compile(r"\x00MATH(?:INLINE|BLOCK)\d+\x00")

# Patterns for code regions that must be shielded from math replacement.
# Fenced code blocks: ``` or ~~~ delimiters, content in between.
_FENCED_CODE_RE = re.compile(
    r"(?m)^( {0,3})(```+|~~~+)(\S*)\n(.*?)\n\1\2\s*$",
    re.DOTALL,
)
# Inline code: `...` (backtick pairs).
_INLINE_CODE_RE = re.compile(r"(?<!\\)(`+)((?:[^`]|\n)+?)(?<!\\)\1")

# Sentinel markers used for temporarily stashing protected code spans.
# Use codepoints from the Unicode Private Use Area (U+E000–U+F8FF) to
# avoid collisions with any real document text.
_CODESPAN_SENTINEL = "CS"
_FENCE_SENTINEL = "FC"


def _escape_math(text: str) -> tuple[str, dict[str, str]]:
    """Temporarily replace math spans with placeholders so mistune won't
    mangle them.  Returns ``(processed_text, placeholder_map).``

    Fenced code blocks and inline code spans are protected first so that
    ``$`` characters inside them (e.g. the assembly ``JNB TF0, $`` idiom)
    are not mistaken for LaTeX math delimiters.
    """
    placeholders: dict[str, str] = {}
    counter = 0

    # -- 1. Stash fenced code blocks and inline code -----------------------
    protected: dict[str, str] = {}
    pc = 0

    def _stash_fence(m: re.Match) -> str:
        nonlocal pc
        key = f"{_FENCE_SENTINEL}{pc}"
        protected[key] = m.group(0)  # preserve the entire block verbatim
        pc += 1
        return key

    def _stash_codespan(m: re.Match) -> str:
        nonlocal pc
        key = f"{_CODESPAN_SENTINEL}{pc}"
        protected[key] = m.group(0)  # preserve backticks + content
        pc += 1
        return key

    text = _FENCED_CODE_RE.sub(_stash_fence, text)
    text = _INLINE_CODE_RE.sub(_stash_codespan, text)

    # -- 2. Replace math spans with placeholders ---------------------------

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

    # -- 3. Restore stashed code regions -----------------------------------
    # Sort by key length descending to prevent short keys from corrupting
    # longer keys that share the same prefix (e.g. "FC1" is a prefix
    # of "FC10" — restoring the former first turns the latter into
    # "<content>0", leaking orphan digits and losing the original content).

    for key, value in sorted(
        protected.items(), key=lambda item: len(item[0]), reverse=True
    ):
        text = text.replace(key, value)

    return text, placeholders


def _restore_text(text: str, placeholders: dict[str, str]) -> str:
    """
    The inverse operation of _escape_math.
    Restore math placeholders in a text string, returning the raw math.
    """
    # Sort by key length descending so that "\x00MATHBLOCK10\x00" is
    # restored before "\x00MATHBLOCK1\x00", avoiding prefix collisions.
    for key, value in sorted(
        placeholders.items(), key=lambda item: len(item[0]), reverse=True
    ):
        text = text.replace(key, value)
    return text


# ---------------------------------------------------------------------------
# Inline children extraction
# ---------------------------------------------------------------------------


def _split_by_placeholders(
    text: str, placeholders: dict[str, str]
) -> list[dict[str, Any]]:
    """Split text on math placeholder markers, producing text + inline_math runs."""
    runs: list[dict[str, Any]] = []
    parts = _PH_ANY.split(text)  # split on placeholders, keeping text parts
    markers = _PH_ANY.findall(text)  # find all placeholders in order

    # parts and markers interleave: part0, marker0, part1, marker1, ...
    for i, part in enumerate(parts):
        if part.strip():  # part is non-empty text
            runs.append({"type": "text", "text": part})
        if i < len(markers):
            marker = markers[i]
            math_content = placeholders.get(
                marker, marker
            )  # Mapping failed, display the placeholder itself
            runs.append({"type": "inline_math", "text": math_content})
    return runs


def _extract_inline_children(
    children: list[dict[str, Any]],
    placeholders: dict[str, str],
) -> list[dict[str, Any]]:
    """Convert mistune inline AST children into our run-dict format.

    Attention: mistune uses ``"raw"`` for text nodes (not ``"text"``).
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

        elif ctype == "strong":  # bold
            inner = _extract_inline_children(child.get("children", []), placeholders)
            result.append({"type": "bold", "children": inner})

        elif ctype == "emphasis":  # italic
            inner = _extract_inline_children(child.get("children", []), placeholders)
            result.append({"type": "italic", "children": inner})

        elif ctype == "codespan":
            raw = child.get("raw", "")
            result.append({"type": "code", "text": raw})

        elif ctype == "link":
            inner = _extract_inline_children(child.get("children", []), placeholders)
            result.append(
                {
                    "type": "link",
                    "url": child.get("link", ""),
                    "children": inner,
                }
            )

        elif ctype == "image":
            # mistune: attrs.url for src, children[text] for alt
            # mistune URL-encodes non-ASCII characters in the path; decode them
            # so the file system can find the actual file.
            src = unquote(child.get("attrs", {}).get("url", ""))
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


# ---------------------------------------------------------------------------
# Block conversion
# ---------------------------------------------------------------------------


def _block_to_dict(
    block: dict[str, Any],
    placeholders: dict[str, str],
) -> dict[str, Any] | None:
    """Convert a single mistune AST block to our intermediate format."""
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

    # --- Table (mistune: table → table_head/table_body → table_row → table_cell) ---
    if btype == "table":
        header_runs: list[list[dict[str, Any]]] = []
        body_rows: list[list[list[dict[str, Any]]]] = []

        for section in block.get("children", []):
            stype = section.get("type", "")
            if stype == "table_head":
                # table_head children are table_cell directly (no table_row)
                for cell in section.get("children", []):
                    runs = _extract_inline_children(
                        cell.get("children", []), placeholders
                    )
                    # If all runs are plain text, flatten to a single string for
                    # backward compatibility with simpler downstream code.
                    if runs and all(r.get("type") == "text" for r in runs):
                        header_runs.append("".join(r.get("text", "") for r in runs))
                    else:
                        header_runs.append(runs)
            elif stype == "table_body":
                # table_body children are table_row → table_cell
                for row in section.get("children", []):
                    row_runs: list[list[dict[str, Any]] | str] = []
                    for cell in row.get("children", []):
                        runs = _extract_inline_children(
                            cell.get("children", []), placeholders
                        )
                        if runs and all(r.get("type") == "text" for r in runs):
                            row_runs.append("".join(r.get("text", "") for r in runs))
                        else:
                            row_runs.append(runs)
                    body_rows.append(row_runs)

        return {
            "type": "table",
            "header": header_runs,
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
                    runs = _extract_inline_children(
                        ic.get("children", []), placeholders
                    )
                    inline_runs.extend(runs)
                elif ic_type == "list":
                    sub = _block_to_dict(ic, placeholders)
                    if sub:
                        sub_lists.append(sub)
                elif ic_type == "paragraph":
                    runs = _extract_inline_children(
                        ic.get("children", []), placeholders
                    )
                    inline_runs.extend(runs)

            items.append(
                {
                    "type": "list_item",
                    "inline_runs": inline_runs,
                    "children": sub_lists,
                }
            )

        return {"type": "list", "ordered": ordered, "items": items}

    # --- Block HTML (may contain <img>) ---
    if btype == "block_html":
        raw = block.get("raw", "")
        img_match = re.search(r'<img[^>]+src=["\']([^"\']+)["\']', raw)
        if img_match:
            return {"type": "image", "src": unquote(img_match.group(1)), "alt": ""}
        return None

    return None


def _flatten_blocks(
    blocks: list[dict[str, Any]],
    placeholders: dict[str, str],
) -> list[dict[str, Any]]:
    """Convert a list of mistune AST blocks into our flat format."""
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
                elif (
                    isinstance(converted, dict) and converted.get("type") == "paragraph"
                ):
                    # Standalone image(s) in a paragraph → image blocks.
                    # Consecutive images on adjacent lines (separated only by
                    # softbreaks / linebreaks) are also promoted so each image
                    # gets full-width, vertical-stack rendering instead of
                    # being squeezed as a narrow inline image.
                    para_children = converted.get("children", [])
                    if para_children and all(
                        c.get("type") in ("image", "softbreak", "linebreak")
                        for c in para_children
                    ):
                        for c in para_children:
                            if c.get("type") == "image":
                                result.append(c)
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
