#!/usr/bin/env python3
"""md2docx — Convert Markdown files to styled .docx documents.

Usage::

    md2docx input.md                  # Uses config.yaml, outputs input.docx
    md2docx input.md -c my_conf.yaml  # Custom config
    md2docx input.md -o output.docx   # Custom output name
"""

from __future__ import annotations

import argparse
import io
import sys
from pathlib import Path

from md2docx.config import AppConfig, load_config
from md2docx.parser import parse_markdown
from md2docx.builder import build_docx


RED = "\033[31m"
YELLOW = "\033[33m"
GREEN = "\033[32m"
RESET = "\033[0m"

def main(argv: list[str] | None = None) -> int:
    """CLI entry point.  Returns 0 on success, 1 on error."""
    try:
        sys.stdout = io.TextIOWrapper(
            sys.stdout.buffer, encoding="utf-8", errors="replace"
        )
    except Exception:
        pass

    parser = argparse.ArgumentParser(
        prog="md2docx",
        description="将 Markdown 文件转换为格式化的 .docx 文档",
    )
    parser.add_argument(
        "input",
        type=str,
        help="输入的 Markdown 文件路径",
    )
    parser.add_argument(
        "-c", "--config",
        type=str,
        default="config.yaml",
        help="YAML 配置文件路径 (默认: config.yaml)",
    )
    parser.add_argument(
        "-o", "--output",
        type=str,
        default=None,
        help="输出 .docx 文件路径 (默认: 与输入文件同目录同名.docx)",
    )

    args = parser.parse_args(argv)

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"{RED}[Error]{RESET} 输入文件不存在: {input_path}]", file=sys.stderr)
        return 1
    if not input_path.suffix.lower() in (".md", ".markdown", ".mdown", ".mkd"):
        print(f"{RED}[Error]{RESET} 输入文件扩展名不是常见的 Markdown 扩展名: {input_path}", file=sys.stderr)
        return 1

    if args.output:
        output_path = Path(args.output)
    else:
        output_path = input_path.with_suffix(".docx")

    config_path = args.config
    if not Path(config_path).exists():
        print(f"{YELLOW}[Warning]{RESET} 配置文件不存在，使用默认配置", file=sys.stderr)
        config_path = None

    try:
        if config_path:
            config = load_config(config_path)
        else:
            config = AppConfig()
    except Exception as exc:
        print(f"{RED}[Error]{RESET} 加载配置文件失败: {exc}", file=sys.stderr)
        return 1

    try:
        md_text = input_path.read_text(encoding="utf-8")
    except Exception as exc:
        print(f"{RED}[Error]{RESET} 读取 Markdown 文件失败: {exc}", file=sys.stderr)
        return 1

    print(f"解析 Markdown: {input_path}")
    try:
        ast = parse_markdown(md_text)
    except Exception as exc:
        print(f"{RED}[Error]{RESET} 解析 Markdown 失败: {exc}", file=sys.stderr)
        return 1
    print(f"  发现 {len(ast)} 个块级元素")

    print(f"生成 DOCX: {output_path}")
    try:
        doc = build_docx(ast, config, input_path, output_path)
    except Exception as exc:
        print(f"{RED}[Error]{RESET} 生成 DOCX 失败: {exc}", file=sys.stderr)
        return 1

    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(output_path))
    except Exception as exc:
        print(f"{RED}[Error]{RESET} 保存 DOCX 失败: {exc}", file=sys.stderr)
        return 1

    print(f"{GREEN}[OK]{RESET} 转换完成: {output_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
