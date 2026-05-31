"""md2docx — Markdown to DOCX converter with YAML-driven styling."""

from md2docx.config import AppConfig, load_config
from md2docx.parser import parse_markdown
from md2docx.builder import build_docx
from md2docx.cli import main

__all__ = ["AppConfig", "load_config", "parse_markdown", "build_docx", "main"]
