"""COM automation for SVG post-processing in generated .docx files.

Uses Windows COM (``win32com.client``) to automate Microsoft Word, find SVG
placeholder paragraphs inserted by the builder, and replace them with actual
SVG images via ``InlineShapes.AddPicture()``.

Requirements
------------
* Windows OS
* Microsoft Word 2016 or later (native SVG support)
* ``pywin32`` package: ``pip install pywin32``

Usage::

    from md2docx.svg_com import process_svg_placeholders
    replaced, failed = process_svg_placeholders(Path("output.docx"))
    print(f"{replaced} replaced, {failed} failed")
"""

from __future__ import annotations

import re
from pathlib import Path

from md2docx.svg_utils import parse_placeholder

# Status text for missing/failed SVGs (matches the existing i18n convention)
_SVG_MISSING = "[SVG文件缺失: {name}]"
_SVG_FAILED = "[SVG插入失败: {name}]"

# Regex to locate a full placeholder string in paragraph text.
# The placeholder format is:  [SVG待处理: ...|path=...|w=...|h=...]
# None of the fields contain ']' so this is safe.
_FIND_PLACEHOLDER_RE = re.compile(r"\[SVG待处理:[^\]]+\]")


def is_com_available() -> bool:
    """Return ``True`` if Word COM automation is available on this system.

    Checks that (a) ``pywin32`` is importable and (b) Word can be started.
    """
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False

    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Quit()
        return True
    except Exception:
        return False


def process_svg_placeholders(docx_path: Path) -> tuple[int, int]:
    """Post-process a .docx file, replacing SVG placeholders with actual SVG images.

    Opens Word via COM, searches for ``[SVG待处理: ...]`` placeholders
    (both block-level and inline), and replaces each one with an
    ``InlineShape`` pointing at the SVG file.  The document is saved in-place.

    Args:
        docx_path: Absolute path to the .docx file to process (modified in-place).

    Returns:
        ``(replaced, failed)`` count tuple.
        If Word COM is completely unavailable, returns ``(0, 0)``.
    """
    try:
        import win32com.client
    except ImportError:
        print("错误: 需要 pywin32 来进行 SVG 后处理。请运行: pip install pywin32")
        return 0, 0

    abs_path = str(docx_path.resolve())

    word = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        try:
            word.DisplayAlerts = False
        except Exception:
            pass
    except Exception:
        print("错误: 无法启动 Microsoft Word。请确认已安装 Word 2016 或更新版本。")
        return 0, 0

    try:
        try:
            doc = word.Documents.Open(abs_path)
        except Exception as exc:
            print(f"错误: 无法打开文档: {abs_path}\n  {exc}")
            print("  请确认文档未被其他程序锁定（关闭 Word 后重试）。")
            return 0, 0

        replaced = 0
        failed = 0

        # Build a list of (paragraph, placeholder_text) pairs.
        # Collect first — modifying paragraphs during iteration is unsafe.
        jobs: list = []
        for i in range(1, doc.Paragraphs.Count + 1):
            para = doc.Paragraphs(i)
            text = para.Range.Text
            m = _FIND_PLACEHOLDER_RE.search(text)
            if m:
                jobs.append((para, m.group(0)))

        if not jobs:
            print("  未发现 SVG 占位符，无需处理。")
            doc.Close(SaveChanges=False)
            return 0, 0

        print(f"  发现 {len(jobs)} 个 SVG 占位符，正在替换...")

        for para, placeholder_text in jobs:
            info = parse_placeholder(placeholder_text)
            if info is None:
                failed += 1
                continue

            svg_path, w_in, h_in = info
            svg_file = Path(svg_path)

            if not svg_file.exists():
                # Replace the placeholder text with a missing-file note
                _replace_text_in_para(
                    para,
                    placeholder_text,
                    _SVG_MISSING.format(name=svg_file.name),
                )
                print(f"  [跳过] SVG 文件不存在: {svg_file.name}")
                failed += 1
                continue

            try:
                # Delete the placeholder text and insert the SVG image at
                # the same position.
                _replace_text_with_svg(
                    para, placeholder_text, str(svg_file), w_in, h_in
                )
                print(f"  [OK] {svg_file.name}")
                replaced += 1
            except Exception as exc:
                _replace_text_in_para(
                    para,
                    placeholder_text,
                    _SVG_FAILED.format(name=svg_file.name),
                )
                print(f"  [失败] {svg_file.name}: {exc}")
                failed += 1

        doc.Save()
        doc.Close()

        return replaced, failed

    finally:
        try:
            word.Quit()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Internal COM helpers
# ---------------------------------------------------------------------------


def _find_text_range(para, search_text: str):
    """Use Word's Find to locate *search_text* within *para*.

    Returns the ``Range`` covering the found text, or ``None`` if not found.
    The search is case-sensitive and literal (no wildcards).
    """
    # Work on a copy of the paragraph range
    rng = para.Range.Duplicate
    rng.Find.Text = search_text
    rng.Find.MatchWildcards = False
    rng.Find.MatchCase = True
    rng.Find.Forward = True
    rng.Find.Wrap = False  # don't wrap beyond the paragraph
    if rng.Find.Execute():
        return rng
    return None


def _replace_text_in_para(para, old_text: str, new_text: str) -> None:
    """Replace *old_text* with *new_text* inside *para* using Word Find."""
    rng = _find_text_range(para, old_text)
    if rng is not None:
        rng.Text = new_text


def _replace_text_with_svg(
    para, old_text: str, svg_abs_path: str, w_in: float, h_in: float
) -> None:
    """Replace *old_text* with an SVG InlineShape inside *para*."""
    rng = _find_text_range(para, old_text)
    if rng is None:
        raise RuntimeError("无法在段落中定位占位符文本")

    # Clear the text first
    rng.Text = ""
    # Insert SVG at the now-empty range
    shape = para.Range.Document.InlineShapes.AddPicture(
        FileName=svg_abs_path,
        LinkToFile=False,
        SaveWithDocument=True,
        Range=rng,
    )
    # Dimensions in points (Word COM convention)
    shape.Width = w_in * 72.0
    shape.Height = h_in * 72.0
