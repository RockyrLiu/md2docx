#!/usr/bin/env python3
"""md2docx — Convert Markdown files to styled .docx documents.

Usage::

    md2docx input.md                  # Uses config.yaml, outputs input.docx
    md2docx input.md -c my_conf.yaml  # Custom config
    md2docx input.md -o output.docx   # Custom output name
    md2docx a.md b.md c.md            # Multi-file, outputs a.docx b.docx c.docx
    md2docx *.md                      # Glob expansion, convert all .md files
    md2docx *.md -o out/              # Output directory for multi-file mode
    md2docx *.md --exclude README.md  # Exclude specific files
    md2docx *.md --exclude "test_*"   # Exclude by glob pattern
    md2docx -v                        # Show version number and exit
    md2docx -ic                       # Write default config.yaml and exit
    md2docx -ic my_conf.yaml          # Write config to a custom path
"""

from __future__ import annotations

import argparse
import io
import sys
from importlib.metadata import version
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


def _expand_inputs(
    inputs: list[str],
    excludes: list[str],
) -> list[Path]:
    """Expand glob patterns in *inputs*, apply *excludes*, return sorted unique paths.

    Each entry in *inputs* may be a literal path or a glob pattern.  Each entry
    in *excludes* may be a literal filename or a glob pattern applied against
    the resolved file name and path.

    Returns a sorted list of unique ``Path`` objects that exist and have a
    recognised Markdown extension.
    """
    md_extensions = {".md", ".markdown", ".mdown", ".mkd"}
    seen: set[Path] = set()

    # -- 1. Expand input patterns -----------------------------------------------
    for pattern in inputs:
        p = Path(pattern)
        # If the pattern contains wildcard characters, glob it
        if any(ch in pattern for ch in ("*", "?", "[")):
            hits = list(Path(".").glob(pattern))
            for hit in hits:
                if hit.is_file() and hit.suffix.lower() in md_extensions:
                    seen.add(hit.resolve())
        elif p.is_file():
            seen.add(p.resolve())
        else:
            print(
                f"{YELLOW}[Warning]{RESET} 跳过不存在的文件: {pattern}",
                file=sys.stderr,
            )

    if not seen:
        return []

    # -- 2. Build exclude set ---------------------------------------------------
    excluded: set[Path] = set()
    for exc_pattern in excludes:
        has_wildcard = any(ch in exc_pattern for ch in ("*", "?", "["))
        for fpath in seen:
            matched = False
            if has_wildcard:
                # Glob pattern: match against filename and full path
                if fpath.match(exc_pattern) or Path(fpath.name).match(exc_pattern):
                    matched = True
            else:
                # Bare name: match filename exactly (case-insensitive on Windows)
                if fpath.name.lower() == Path(exc_pattern).name.lower():
                    matched = True
                # Also allow matching by full path (e.g. "subdir/file.md")
                if str(fpath).endswith(exc_pattern) or fpath.match(exc_pattern):
                    matched = True
            if matched:
                excluded.add(fpath)

    # -- 3. Filter and sort -----------------------------------------------------
    result = sorted(seen - excluded)
    if excluded:
        for exc_f in sorted(excluded):
            print(f"{YELLOW}[排除]{RESET} {exc_f.name}", file=sys.stderr)
    return result


def _convert_one(
    input_path: Path,
    output_path: Path,
    config: AppConfig,
) -> int:
    """Convert a single Markdown file to DOCX.  Returns 0 on success, 1 on error."""
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


def main(argv: list[str] | None = None) -> int:
    """CLI entry point.  Returns 0 on success, 1 on error."""

    parser = argparse.ArgumentParser(
        prog="md2docx",
        description="将 Markdown 文件转换为格式化的 .docx 文档",
    )
    parser.add_argument(
        "input",
        nargs="*",
        type=str,
        default=None,
        help="输入的 Markdown 文件路径 (支持通配符，如 *.md)",
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
        help="输出 .docx 文件/目录路径 "
             "(单文件: 指定输出文件; 多文件: 指定输出目录)",
    )
    parser.add_argument(
        "--exclude",
        action="append",
        default=None,
        type=str,
        dest="exclude",
        help="排除指定的 Markdown 文件，支持通配符 (可重复使用)",
    )
    parser.add_argument(
        "-v", "--version",
        action="version",
        version=f"md2docx {version('md2docx')}",
        help="显示版本号并退出",
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

    # Normal conversion mode: input file(s) required
    if not args.input:
        parser.print_help()
        print(f"\n{YELLOW}[Hint]{RESET} 使用 --init-config 生成默认配置文件模板", file=sys.stderr)
        return 1

    # -- Expand inputs (glob + exclude) -----------------------------------------
    excludes: list[str] = args.exclude if args.exclude else []
    input_paths = _expand_inputs(args.input, excludes)

    if not input_paths:
        print(f"{RED}[Error]{RESET} 没有找到有效的 Markdown 输入文件", file=sys.stderr)
        return 1

    print(f"找到 {len(input_paths)} 个 Markdown 文件待转换")

    # -- Validate extension -----------------------------------------------------
    md_extensions = {".md", ".markdown", ".mdown", ".mkd"}
    for p in input_paths:
        if p.suffix.lower() not in md_extensions:
            print(
                f"{RED}[Error]{RESET} 输入文件扩展名不是常见的 Markdown 扩展名: {p}",
                file=sys.stderr,
            )
            return 1

    # -- Load config once -------------------------------------------------------
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

    # -- Determine output mode --------------------------------------------------
    multi_file = len(input_paths) > 1
    output_dir: Path | None = None

    if args.output:
        output_arg = Path(args.output)
        if multi_file:
            # Multiple inputs: -o specifies an output directory
            output_dir = output_arg
        else:
            # Single input: -o specifies the output file (existing behaviour)
            output_dir = None
    else:
        output_dir = None

    # -- Convert each file ------------------------------------------------------
    success_count = 0
    fail_count = 0

    for i, input_path in enumerate(input_paths):
        if i > 0:
            print()
        # Determine output path for this file
        if multi_file:
            if output_dir:
                out = output_dir / input_path.with_suffix(".docx").name
            else:
                out = input_path.with_suffix(".docx")
        else:
            if args.output:
                out = Path(args.output)
            else:
                out = input_path.with_suffix(".docx")

        ret = _convert_one(input_path, out, config)
        if ret == 0:
            success_count += 1
        else:
            fail_count += 1

    # -- Summary ----------------------------------------------------------------
    if multi_file:
        print()
        if fail_count == 0:
            print(f"{GREEN}[OK]{RESET} 全部转换完成: {success_count} 个文件")
        else:
            print(
                f"{YELLOW}[完成]{RESET} "
                f"{GREEN}{success_count}{RESET} 个成功, "
                f"{RED}{fail_count}{RESET} 个失败"
            )

    return 0 if fail_count == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
