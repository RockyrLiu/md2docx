#!/usr/bin/env python3
"""md2docx — Convert Markdown files to styled .docx documents.

Usage::

    md2docx input.md                  # Uses config.yaml, outputs input.docx
    md2docx input.md -c my_conf.yaml  # Custom config
    md2docx input.md -o output.docx   # Custom output name
    md2docx -ic                       # Write default config.yaml and exit
    md2docx -ic my_conf.yaml          # Write config to a custom path
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

# Path to the default config template shipped inside the package
_PACKAGE_DIR = Path(__file__).resolve().parent
_DEFAULT_CONFIG_TEMPLATE = _PACKAGE_DIR / "default_config.yaml"


def _generate_config(target: Path) -> int:
    """Copy the shipped default_config.yaml to *target* as a starter template.

    Returns 0 on success, 1 if the file already exists (refuses to overwrite).
    """
    if target.exists():
        print(
            f"{YELLOW}[Warning]{RESET} 目标文件已存在: {target}\n"
            f"  如需重新生成，请先删除或重命名现有文件。",
            file=sys.stderr,
        )
        return 1

    content = _DEFAULT_CONFIG_TEMPLATE.read_text(encoding="utf-8")
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(content, encoding="utf-8")
    print(f"{GREEN}[OK]{RESET} 配置文件模板已生成: {target}")
    return 0


def main(argv: list[str] | None = None) -> int:
    """CLI entry point.  Returns 0 on success, 1 on error."""

    parser = argparse.ArgumentParser(
        prog="md2docx",
        description="将 Markdown 文件转换为格式化的 .docx 文档",
    )
    parser.add_argument(
        "input",
        nargs="?",
        type=str,
        default=None,
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
    parser.add_argument(
        "-ic", "--init-config",
        nargs="?",
        const="config.yaml",
        default=None,
        type=str,
        dest="init_config",
        help="生成默认配置文件模板并退出 (可选指定输出路径，默认: config.yaml)",
    )

    args = parser.parse_args(argv)

    # --init-config mode: write template and exit (no input file needed)
    if args.init_config is not None:
        return _generate_config(Path(args.init_config))

    # Wrap stdout for UTF-8 output (must happen after --init-config check,
    # because wrapping twice in the same process breaks the underlying buffer).
    try:
        sys.stdout = io.TextIOWrapper(
            sys.stdout.buffer, encoding="utf-8", errors="replace",
            line_buffering=True,
        )
    except Exception:
        pass

    # Normal conversion mode: input file is required
    if args.input is None:
        parser.print_help()
        print(f"\n{YELLOW}[Hint]{RESET} 使用 --init-config 生成默认配置文件模板", file=sys.stderr)
        return 1

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"{RED}[Error]{RESET} 输入文件不存在: {input_path}", file=sys.stderr)
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
        doc = build_docx(ast, config, input_path)
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
