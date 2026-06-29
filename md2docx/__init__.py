"""md2docx — Markdown to DOCX converter with YAML-driven styling."""

from md2docx.config import AppConfig, load_config
from md2docx.parser import parse_markdown
from md2docx.builder import build_docx
from md2docx.cli import main
from md2docx.svg_utils import (
    is_svg,
    parse_svg_dimensions,
    make_placeholder,
    parse_placeholder,
)
from md2docx.svg_com import process_svg_placeholders, is_com_available

__all__ = [
    "AppConfig",
    "load_config",
    "parse_markdown",
    "build_docx",
    "main",
    "is_svg",
    "parse_svg_dimensions",
    "make_placeholder",
    "parse_placeholder",
    "process_svg_placeholders",
    "is_com_available",
]
