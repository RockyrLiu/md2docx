"""SVG utility functions — dimension parsing and placeholder generation/parsing.

Used by both the builder (Phase 1: insert placeholders for SVG images) and
the COM post-processor (Phase 2: find placeholders and replace with SVGs).
"""

from __future__ import annotations

import re
from pathlib import Path

from lxml import etree

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

SVG_DEFAULT_WIDTH_IN = 3.0  # inches
SVG_DEFAULT_HEIGHT_IN = 3.0  # inches
SVG_DPI = 96.0

# Placeholder format:
#   [SVG待处理: filename.svg (W"xH")|path=absolute_path|w=W|h=H]
# The | character is illegal in Windows paths, so it is a safe delimiter.
_PLACEHOLDER_RE = re.compile(
    r"\[SVG待处理:\s*(.+?)\s*\(([\d.]+)\"x([\d.]+)\"\)"
    r"\|path=(.+?)\|w=([\d.]+)\|h=([\d.]+)\]"
)

# SVG unit → inches conversion factor
_UNIT_TO_INCHES: dict[str, float] = {
    "in": 1.0,
    "px": 1.0 / 96.0,
    "pt": 1.0 / 72.0,
    "pc": 12.0 / 72.0,
    "mm": 1.0 / 25.4,
    "cm": 1.0 / 2.54,
    "em": 16.0 / 96.0,  # approximate — depends on font context
    "ex": 8.0 / 96.0,  # approximate
}

# Regex for parsing an SVG length: number + optional unit
_UNIT_RE = re.compile(r"^\s*([\d.]+)\s*(px|pt|pc|mm|cm|in|em|ex|%)\s*$", re.IGNORECASE)


# ---------------------------------------------------------------------------
# Public helpers
# ---------------------------------------------------------------------------


def is_svg(path: str | Path) -> bool:
    """Return True if *path* has a ``.svg`` extension (case-insensitive)."""
    return Path(path).suffix.lower() == ".svg"


def parse_svg_dimensions(svg_path: str | Path) -> tuple[float, float]:
    """Parse the natural width and height (in inches) from an SVG file.

    Priority:
    1. ``viewBox`` attribute → use its width and height (SVG user units at
       96 DPI, which means 1 unit = 1/96 inch)
    2. ``width`` / ``height`` attributes with unit conversion
    3. Default (3.0" × 3.0")

    Returns ``(width_inches, height_inches)``.
    """
    try:
        tree = etree.parse(str(svg_path))
        root = tree.getroot()
        attrs = root.attrib
    except Exception:
        return SVG_DEFAULT_WIDTH_IN, SVG_DEFAULT_HEIGHT_IN

    width_in = SVG_DEFAULT_WIDTH_IN
    height_in = SVG_DEFAULT_HEIGHT_IN

    # --- viewBox ---
    viewbox = attrs.get("viewBox", "")
    if viewbox:
        parts = viewbox.strip().split()
        if len(parts) >= 4:
            try:
                vb_w = float(parts[2])
                vb_h = float(parts[3])
                if vb_w > 0 and vb_h > 0:
                    width_in = vb_w / SVG_DPI
                    height_in = vb_h / SVG_DPI
                    return (width_in, height_in)
            except ValueError:
                pass

    # --- width / height ---
    w_str = attrs.get("width", "")
    h_str = attrs.get("height", "")

    w_parsed = _parse_svg_unit(w_str) if w_str else None
    h_parsed = _parse_svg_unit(h_str) if h_str else None

    if w_parsed is not None and h_parsed is not None:
        return (w_parsed, h_parsed)

    if w_parsed is not None and h_parsed is None:
        # Only width given — assume square aspect ratio
        return (w_parsed, w_parsed)

    if h_parsed is not None and w_parsed is None:
        return (h_parsed, h_parsed)

    return (width_in, height_in)


def make_placeholder(img_path: str | Path, width_in: float, height_in: float) -> str:
    """Build a COM-detectable placeholder string for an SVG image.

    The format is::

        [SVG待处理: filename.svg (W"xH")|path=absolute_path|w=W|h=H]

    The placeholder is human-readable (shows filename and intended size) and
    can be reliably parsed back by :func:`parse_placeholder`.
    """
    p = Path(img_path)
    filename = p.name
    abs_path = str(p.resolve())
    return (
        f'[SVG待处理: {filename} ({width_in:.1f}"x{height_in:.1f}")'
        f"|path={abs_path}|w={width_in:.4f}|h={height_in:.4f}]"
    )


def parse_placeholder(text: str) -> tuple[str, float, float] | None:
    """Try to parse a placeholder string from *text*.

    Returns ``(absolute_path, width_inches, height_inches)`` if *text*
    contains a valid placeholder, or ``None`` otherwise.
    """
    m = _PLACEHOLDER_RE.search(text)
    if not m:
        return None
    try:
        path = m.group(4).strip()
        w = float(m.group(5))
        h = float(m.group(6))
        return (path, w, h)
    except (ValueError, IndexError):
        return None


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------


def _parse_svg_unit(value_str: str) -> float | None:
    """Convert an SVG length string (``"100px"``, ``"5cm"``, ``"2in"``) to
    inches.  Returns ``None`` for percentage values or unparseable input.
    """
    if not value_str:
        return None

    # Unitless numbers are treated as px per SVG spec
    try:
        num = float(value_str)
        return num / SVG_DPI
    except ValueError:
        pass

    m = _UNIT_RE.match(value_str)
    if not m:
        return None

    num = float(m.group(1))
    unit = m.group(2).lower()

    if unit == "%":
        return None  # percentage cannot be resolved without context

    factor = _UNIT_TO_INCHES.get(unit, 1.0 / SVG_DPI)
    return num * factor
