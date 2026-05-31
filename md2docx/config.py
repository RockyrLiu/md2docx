"""Configuration management — YAML loading, validation, and defaults."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Any

import yaml


# =============================================================================
# Style dataclasses
# =============================================================================


@dataclass
class TextStyle:
    """Base text style configuration."""

    font_name: str = "宋体"
    font_name_ascii: str = "Times New Roman"
    font_size: int = 12  # pt
    line_spacing: float = 1.5
    bold: bool = False
    italic: bool = False
    alignment: str = "left"  # left | center | right | justify
    space_before: int = 0  # pt
    space_after: int = 0  # pt
    first_line_indent: int = 0  # 首行缩进字符数
    color: str | None = None  # RGB hex, None = auto


@dataclass
class HeadingStyle:
    """Heading style configuration for a single level."""

    font_name: str = "黑体"
    font_name_ascii: str = "Times New Roman"
    font_size: int = 14
    bold: bool = True
    italic: bool = False
    alignment: str = "left"
    space_before: int = 6
    space_after: int = 6
    color: str | None = None


@dataclass
class TableStyle:
    """Table text style configuration."""

    font_name: str = "宋体"
    font_name_ascii: str = "Times New Roman"
    font_size: int = 10
    bold: bool = False
    header_bold: bool = True
    alignment: str = "center"
    line_spacing: float = 1.0


@dataclass
class ListStyle:
    """List style configuration."""

    font_name: str = "宋体"
    font_name_ascii: str = "Times New Roman"
    font_size: int = 12
    line_spacing: float = 1.5
    indent: float = 0.5  # inches


@dataclass
class CodeStyle:
    """Code block style configuration."""

    font_name: str = "Consolas"
    font_name_ascii: str = "Consolas"
    font_size: int = 10
    line_spacing: float = 1.0
    background_color: str = "F2F2F2"


@dataclass
class BlockquoteStyle:
    """Blockquote style configuration."""

    font_name: str = "楷体"
    font_name_ascii: str = "Times New Roman"
    font_size: int = 12
    italic: bool = True
    line_spacing: float = 1.5
    left_indent: float = 0.5
    border_left_color: str = "999999"


@dataclass
class StyleConfig:
    """Aggregate style configuration."""

    body: TextStyle = field(default_factory=TextStyle)
    headings: dict[str, HeadingStyle] = field(default_factory=dict)
    table: TableStyle = field(default_factory=TableStyle)
    list: ListStyle = field(default_factory=ListStyle)
    code: CodeStyle = field(default_factory=CodeStyle)
    blockquote: BlockquoteStyle = field(default_factory=BlockquoteStyle)


# =============================================================================
# Cover / TOC / Page dataclasses
# =============================================================================


@dataclass
class CoverConfig:
    """Cover page configuration."""

    enabled: bool = True
    title: str = "报告标题"
    subtitle: str = ""
    author: str = ""
    class_info: str = ""
    teacher: str = ""
    department: str = ""
    date: str = ""


@dataclass
class TocConfig:
    """Table of contents configuration."""

    enabled: bool = True
    level: int = 3  # 1-6


@dataclass
class PageConfig:
    """Page layout configuration."""

    size: str = "A4"  # A4 | Letter
    margin_top: float = 2.54  # cm
    margin_bottom: float = 2.54
    margin_left: float = 3.18
    margin_right: float = 3.18


# =============================================================================
# Root config
# =============================================================================


@dataclass
class AppConfig:
    """Root application configuration."""

    cover: CoverConfig = field(default_factory=CoverConfig)
    toc: TocConfig = field(default_factory=TocConfig)
    styles: StyleConfig = field(default_factory=StyleConfig)
    page: PageConfig = field(default_factory=PageConfig)


# =============================================================================
# Built-in heading defaults
# =============================================================================

HEADING_DEFAULTS: dict[str, HeadingStyle] = {
    "h1": HeadingStyle(font_size=22, bold=True, alignment="center", space_before=12, space_after=12),
    "h2": HeadingStyle(font_size=16, bold=True, space_before=6, space_after=6),
    "h3": HeadingStyle(font_size=14, bold=True, space_before=6, space_after=6),
    "h4": HeadingStyle(font_size=12, bold=True, space_before=3, space_after=3),
    "h5": HeadingStyle(font_size=12, bold=True, space_before=3, space_after=3),
    "h6": HeadingStyle(font_size=12, bold=True, space_before=3, space_after=3),
}


# =============================================================================
# YAML loading helpers
# =============================================================================


def _deep_merge(base: dict, override: dict) -> dict:
    """Recursively merge override into base dict."""
    merged = dict(base)
    for key, value in override.items():
        if key in merged and isinstance(merged[key], dict) and isinstance(value, dict):
            merged[key] = _deep_merge(merged[key], value)
        else:
            merged[key] = value
    return merged


def _dict_to_dataclass(cls: type, data: dict[str, Any]) -> Any:
    """Convert a dict to a dataclass instance, dropping unknown keys."""
    field_names = {f.name for f in cls.__dataclass_fields__.values()}  # type: ignore[attr-defined]
    kwargs = {k: v for k, v in data.items() if k in field_names}
    return cls(**kwargs)


def _parse_headings(data: dict[str, Any]) -> dict[str, HeadingStyle]:
    """Parse heading styles with built-in defaults."""
    result: dict[str, HeadingStyle] = {}
    for level in ("h1", "h2", "h3", "h4", "h5", "h6"):
        default = HEADING_DEFAULTS[level]
        if level in data:
            merged = _deep_merge(default.__dict__ if hasattr(default, "__dict__") else {}, data[level])
            result[level] = _dict_to_dataclass(HeadingStyle, merged)
        else:
            result[level] = default
    return result


def load_config(path: str | Path) -> AppConfig:
    """Load configuration from a YAML file, filling in defaults for missing keys.

    Args:
        path: Path to the YAML configuration file.

    Returns:
        Fully-populated AppConfig instance.

    Raises:
        FileNotFoundError: If the config file does not exist.
        yaml.YAMLError: If the config file contains invalid YAML.
    """
    config_path = Path(path)
    if not config_path.exists():
        raise FileNotFoundError(f"配置文件不存在: {config_path}")

    with open(config_path, encoding="utf-8") as f:
        raw: dict[str, Any] = yaml.safe_load(f) or {}

    # --- Cover ---
    cover_data = raw.get("cover", {})
    cover = _dict_to_dataclass(CoverConfig, cover_data)
    if not cover.date:
        cover.date = date.today().strftime("%Y-%m-%d")

    # --- TOC ---
    toc_data = raw.get("toc", {})
    toc = _dict_to_dataclass(TocConfig, toc_data)

    # --- Styles ---
    styles_data = raw.get("styles", {})

    body_data = styles_data.get("body", {})
    body = _dict_to_dataclass(TextStyle, body_data)

    headings_data = styles_data.get("headings", {})
    headings = _parse_headings(headings_data)

    table_data = styles_data.get("table", {})
    table = _dict_to_dataclass(TableStyle, table_data)

    list_data = styles_data.get("list", {})
    list_style = _dict_to_dataclass(ListStyle, list_data)

    code_data = styles_data.get("code", {})
    code = _dict_to_dataclass(CodeStyle, code_data)

    bq_data = styles_data.get("blockquote", {})
    blockquote = _dict_to_dataclass(BlockquoteStyle, bq_data)

    styles = StyleConfig(
        body=body,
        headings=headings,
        table=table,
        list=list_style,
        code=code,
        blockquote=blockquote,
    )

    # --- Page ---
    page_data = raw.get("page", {})
    page = _dict_to_dataclass(PageConfig, page_data)

    return AppConfig(cover=cover, toc=toc, styles=styles, page=page)
