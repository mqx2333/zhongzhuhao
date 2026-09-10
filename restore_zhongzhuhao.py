#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
restore_zhongzhuhao.py
======================

辅助 pandoc 修复 Word 转换到 Markdown 时丢失的两种字符标注：

    1. 着重号（字下加点）：docx 中的 ``<w:em w:val="dot"/>``
    2. 字体颜色：        docx 中的 ``<w:color w:val="RRGGBB"/>``

Pandoc 转 Markdown 时会忽略二者，本程序读回 docx 的标注，在 Markdown 中
用内联 HTML ``<span>`` 补回（Markdown 支持内联 HTML，pandoc 再转 HTML 时原样透传），
可选注入 ``text-emphasis`` 样式渲染字下加点。

    <span class="zhongzhuhao">着重文字</span>
    <span style="color:#FF0000">红色文字</span>
    <span class="zhongzhuhao" style="color:#0000FF">又蓝又着重</span>

用法（CLI）：
    python restore_zhongzhuhao.py --docx src.docx --markdown pandoc输出.md -o 修复后.md
    python restore_zhongzhuhao.py --docx src.docx --run-pandoc -o 修复后.md

用法（GUI）：
    python gui.py
"""

from __future__ import annotations

import argparse
import re
import subprocess
import sys
import zipfile
from dataclasses import dataclass, replace
from pathlib import Path

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
DEFAULT_EMPHASIS_CLASS = "zhongzhuhao"

WS_RE = re.compile(r"\s+")
BLANK_LINE_RE = re.compile(r"\n[ \t]*\n")
HEX6_RE = re.compile(r"^[0-9a-fA-F]{6}$")
HEX3_RE = re.compile(r"^[0-9a-fA-F]{3}$")


def wtag(name: str) -> str:
    return "{%s}%s" % (W_NS, name)


def normalize(s: str) -> str:
    return WS_RE.sub(" ", s).strip()


def default_css(emphasis_class: str = DEFAULT_EMPHASIS_CLASS) -> str:
    return (
        ".%s {"
        "text-emphasis-style: dot;"
        "text-emphasis-position: under;"
        "text-emphasis-color: currentColor;"
        "}" % emphasis_class
    )


# --------------------------------------------------------------------------- #
# 数据模型：一处字符标注
# --------------------------------------------------------------------------- #
@dataclass
class Annotation:
    """一处字符的标注。新增标注类型时在此加字段即可。"""

    emphasis: bool = False          # 着重号（字下加点）
    color: str | None = None        # 形如 "#RRGGBB"，None 表示无颜色

    def is_empty(self) -> bool:
        return not self.emphasis and self.color is None


@dataclass
class Segment:
    """连续且标注相同的一段文字。"""

    text: str
    annot: Annotation


@dataclass
class Options:
    """处理选项（CLI 与 GUI 共用）。"""

    emphasis: bool = True
    color: bool = True
    emphasis_class: str = DEFAULT_EMPHASIS_CLASS
    inject_style: bool = True
    css: str | None = None


# --------------------------------------------------------------------------- #
# 1. 从 docx 提取标注（着重号 / 字体颜色）
# --------------------------------------------------------------------------- #
def _run_emphasis(rpr) -> bool:
    if rpr is None:
        return False
    em = rpr.find(wtag("em"))
    return em is not None and em.get(wtag("val")) not in (None, "none")


def _run_color(rpr) -> str | None:
    """读取 w:color；返回 "#RRGGBB" 或 None（auto / 主题色 / 非法值）。"""
    if rpr is None:
        return None
    color = rpr.find(wtag("color"))
    if color is None:
        return None
    val = color.get(wtag("val"))
    if not val or val.lower() == "auto":
        return None
    val = val.lstrip("#")
    if HEX6_RE.match(val):
        return "#" + val.upper()
    if HEX3_RE.match(val):
        return "#" + "".join(c * 2 for c in val).upper()
    return None


def extract_segments(
    docx_path: str | Path,
    want_emphasis: bool = True,
    want_color: bool = True,
) -> list[Segment]:
    """按文档顺序提取所有带标注的文字段。

    同段落内相邻、且标注完全相同的 run 会合并为一段。
    """
    docx_path = Path(docx_path)
    with zipfile.ZipFile(docx_path) as zf:
        try:
            xml = zf.read("word/document.xml")
        except KeyError:
            raise SystemExit("不是有效的 .docx 文件（缺少 word/document.xml）。")

    root = etree.fromstring(xml)
    body = root.find(wtag("body"))
    if body is None:
        return []

    segments: list[Segment] = []
    for para in body.iter(wtag("p")):
        para_segments: list[Segment] = []
        cur: Segment | None = None
        for run in para.iter(wtag("r")):
            text = "".join((t.text or "") for t in run.iter(wtag("t")))
            rpr = run.find(wtag("rPr"))
            annot = Annotation(
                emphasis=_run_emphasis(rpr) if want_emphasis else False,
                color=_run_color(rpr) if want_color else None,
            )
            if annot.is_empty():
                if cur is not None:
                    para_segments.append(cur)
                    cur = None
                continue
            if cur is not None and cur.annot == annot:
                cur.text += text
            else:
                if cur is not None:
                    para_segments.append(cur)
                cur = Segment(text, annot)
        if cur is not None:
            para_segments.append(cur)

        for seg in para_segments:
            seg.text = normalize(seg.text)
            if seg.text:
                segments.append(seg)
    return segments


# --------------------------------------------------------------------------- #
# 2. Markdown 文本匹配（跳过代码块 / 行内代码 / 链接）
# --------------------------------------------------------------------------- #
def mark_fenced_code(lines: list[str]) -> list[bool]:
    """标记围栏代码块（``` 或 ~~~）的行。"""
    is_code = [False] * len(lines)
    i = 0
    while i < len(lines):
        stripped = lines[i].lstrip()
        if stripped.startswith("```") or stripped.startswith("~~~"):
            fence = stripped[:3]
            is_code[i] = True
            j = i + 1
            while j < len(lines) and not lines[j].lstrip().startswith(fence):
                is_code[j] = True
                j += 1
            if j < len(lines):
                is_code[j] = True
            i = j + 1
        else:
            i += 1
    return is_code


def protected_ranges(line: str) -> list[tuple[int, int]]:
    """返回一行中不应被包裹的区间：行内代码、链接/图片。"""
    ranges: list[tuple[int, int]] = []
    n = len(line)
    i = 0
    while i < n:
        ch = line[i]
        if ch == "`":
            j = i + 1
            while j < n and line[j] == "`":
                j += 1
            ticks = j - i
            if "`" * ticks in line[j:]:
                k = line.index("`" * ticks, j)
                ranges.append((i, k + ticks))
                i = k + ticks
            else:
                ranges.append((i, n))
                break
            continue
        if ch in "[!":
            closeb = line.find("]", i + 1)
            if closeb > 0 and closeb + 1 < n and line[closeb + 1] == "(":
                endp = line.find(")", closeb + 2)
                if endp >= 0:
                    ranges.append((i, endp + 1))
                    i = endp + 1
                    continue
        i += 1

    if not ranges:
        return []
    merged = sorted(ranges)
    out = [list(merged[0])]
    for s, e in merged[1:]:
        if s <= out[-1][1]:
            out[-1][1] = max(out[-1][1], e)
        else:
            out.append([s, e])
    return [tuple(x) for x in out]


def compute_protected(markdown: str) -> list[tuple[int, int]]:
    """在整个 Markdown 上计算受保护区间的绝对偏移。"""
    lines = markdown.split("\n")
    is_code = mark_fenced_code(lines)
    protected: list[tuple[int, int]] = []
    offset = 0
    for li, line in enumerate(lines):
        if is_code[li]:
            protected.append((offset, offset + len(line)))
        else:
            for s, e in protected_ranges(line):
                protected.append((offset + s, offset + e))
        offset += len(line) + 1
    return protected


def _overlaps_simple(s: int, e: int, ranges) -> bool:
    for rs, re_ in ranges:
        if rs < e and s < re_:
            return True
    return False


def _pattern_for(phrase: str) -> re.Pattern:
    """把规范化的短语编译成正则：短语中的空格可匹配软换行/多空白。"""
    parts = [re.escape(p) for p in phrase.split(" ")]
    return re.compile(r"[ \t]*(?:\n[ \t]*)?".join(parts))


def _search(compiled: re.Pattern, flat: str, start: int, protected) -> re.Match | None:
    for m in compiled.finditer(flat, start):
        s, e = m.start(), m.end()
        if BLANK_LINE_RE.search(flat[s:e]):
            continue
        if not _overlaps_simple(s, e, protected):
            return m
    return None


def _span_open(annot: Annotation, emphasis_class: str) -> str:
    attrs = []
    if annot.emphasis:
        attrs.append('class="%s"' % emphasis_class)
    if annot.color:
        attrs.append('style="color:%s"' % annot.color)
    return "<span %s>" % " ".join(attrs)


def _shift_protected(protected, s: int, e: int, delta: int):
    out = []
    for ps, pe in protected:
        if pe <= s:
            out.append((ps, pe))
        elif ps >= e:
            out.append((ps + delta, pe + delta))
        else:
            out.append((ps, pe))
    return out


def wrap_segments(
    markdown: str,
    segments: list[Segment],
    emphasis_class: str = DEFAULT_EMPHASIS_CLASS,
    dry_run: bool = False,
) -> tuple[str, dict]:
    """依文档顺序把各标注段在 Markdown 中包裹为内联 <span>。"""
    flat = markdown
    protected = compute_protected(flat)
    pos = 0
    wrapped = 0
    missed: list[str] = []
    by_type = {"emphasis": 0, "color": 0}

    for seg in segments:
        phrase = seg.text
        if not phrase:
            continue

        if dry_run:
            m = _search(_pattern_for(phrase), flat, 0, protected)
            if m is not None:
                print("  [命中] %s" % phrase)
                wrapped += 1
                if seg.annot.emphasis:
                    by_type["emphasis"] += 1
                if seg.annot.color:
                    by_type["color"] += 1
            else:
                print("  [未命中] %s" % phrase)
                missed.append(phrase)
            continue

        m = _search(_pattern_for(phrase), flat, pos, protected)
        if m is None:
            missed.append(phrase)
            print("  [警告] 未在 Markdown 中找到: %s" % phrase, file=sys.stderr)
            continue

        s, e = m.start(), m.end()
        open_tag = _span_open(seg.annot, emphasis_class)
        close_tag = "</span>"
        flat = flat[:s] + open_tag + flat[s:e] + close_tag + flat[e:]
        delta = len(open_tag) + len(close_tag)
        pos = e + delta
        protected = _shift_protected(protected, s, e, delta)
        wrapped += 1
        if seg.annot.emphasis:
            by_type["emphasis"] += 1
        if seg.annot.color:
            by_type["color"] += 1

    return flat, {"wrapped": wrapped, "missed": missed, "by_type": by_type}


# --------------------------------------------------------------------------- #
# 3. 高层处理入口（CLI 与 GUI 共用）
# --------------------------------------------------------------------------- #
def process_markdown(
    docx_path: str | Path,
    markdown: str,
    options: Options,
    dry_run: bool = False,
) -> tuple[str, dict]:
    """完整处理：提取标注 -> 包裹 -> 注入样式。返回 (最终 Markdown, 统计)。"""
    segments = extract_segments(docx_path, options.emphasis, options.color)
    new_md, stats = wrap_segments(
        markdown, segments, options.emphasis_class, dry_run=dry_run
    )
    if not dry_run and options.inject_style and stats["wrapped"] > 0:
        css = options.css or default_css(options.emphasis_class)
        new_md = "<style>\n%s\n</style>\n\n%s" % (css, new_md)
    return new_md, stats


def run_pandoc(docx_path, pandoc_cmd: str = "pandoc", fmt: str = "markdown") -> str:
    """调用 pandoc 把 docx 转成文本；markdown 输出时禁用换行以免拆散标注段。"""
    cmd = (pandoc_cmd or "pandoc").split()
    cmd += [str(docx_path), "-t", fmt]
    if "markdown" in fmt:
        cmd.append("--wrap=none")
    proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
    if proc.returncode != 0:
        raise SystemExit("pandoc 调用失败:\n%s" % proc.stderr)
    return proc.stdout


def markdown_to_html(markdown: str, pandoc_cmd: str = "pandoc") -> str:
    """把（已修复的）Markdown 转成独立 HTML 文档。"""
    cmd = (pandoc_cmd or "pandoc").split()
    cmd += ["-f", "markdown", "-t", "html", "-s", "--wrap=none", "--metadata", "title=文档"]
    proc = subprocess.run(
        cmd, input=markdown, capture_output=True, text=True, encoding="utf-8"
    )
    if proc.returncode != 0:
        raise SystemExit("pandoc 转换 HTML 失败:\n%s" % proc.stderr)
    return proc.stdout


def inject_html_style(html: str, css: str) -> str:
    """把样式插入 HTML 的 <head>（无 head 则插到最前）。"""
    block = "<style>\n%s\n</style>\n" % css
    lower = html.lower()
    idx = lower.find("</head>")
    if idx >= 0:
        return html[:idx] + block + html[idx:]
    idx = lower.find("<body")
    if idx >= 0:
        return html[:idx] + block + html[idx:]
    return block + html


def process(
    docx_path: str | Path,
    source_markdown: str,
    options: Options,
    output_format: str = "markdown",
    pandoc_bin: str = "pandoc",
    dry_run: bool = False,
) -> tuple[str, dict]:
    """最高层入口：按 output_format 产出 'markdown' 或 'html'。"""
    if output_format == "html":
        md_opts = replace(options, inject_style=False)
        new_md, stats = process_markdown(docx_path, source_markdown, md_opts, dry_run=dry_run)
        if dry_run:
            return new_md, stats
        html = markdown_to_html(new_md, pandoc_bin)
        if options.inject_style and stats["wrapped"] > 0:
            css = options.css or default_css(options.emphasis_class)
            html = inject_html_style(html, css)
        return html, stats
    return process_markdown(docx_path, source_markdown, options, dry_run=dry_run)


# --------------------------------------------------------------------------- #
# 4. CLI
# --------------------------------------------------------------------------- #
def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="辅助 pandoc：把 docx 中被忽略的着重号与字体颜色以内联 HTML 补回 Markdown / HTML。"
    )
    parser.add_argument("--docx", required=True, help="源 Word .docx 文件")
    parser.add_argument("--markdown", help="pandoc 生成的 Markdown 文件")
    parser.add_argument("-o", "--out", help="输出文件；省略则写回 --markdown / 同名")
    parser.add_argument("--format", choices=["markdown", "html"], default="markdown",
                        help="输出格式（默认 markdown；html 会用 pandoc 转成独立网页）")
    parser.add_argument("--run-pandoc", dest="pandoc", nargs="?", const="pandoc",
                        metavar="PANDOC", help="先调用 pandoc 把 docx 转成 markdown 再修复")
    parser.add_argument("--pandoc-bin", default="pandoc", help="pandoc 可执行文件（html 转换用）")
    parser.add_argument("--to", default="markdown", help="pandoc 输出格式（默认 markdown）")
    parser.add_argument("--no-emphasis", action="store_true", help="不处理着重号")
    parser.add_argument("--no-color", action="store_true", help="不处理字体颜色")
    parser.add_argument("--emphasis-class", default=DEFAULT_EMPHASIS_CLASS,
                        help="着重号 span 的 class 名（默认 %s）" % DEFAULT_EMPHASIS_CLASS)
    parser.add_argument("--css", default=None, help="自定义 .zhongzhuhao 样式")
    parser.add_argument("--no-style", action="store_true", help="不注入样式块")
    parser.add_argument("--dry-run", action="store_true", help="只打印将处理的短语，不写文件")
    parser.add_argument("--encoding", default="utf-8", help="文件读写编码（默认 utf-8）")
    args = parser.parse_args(argv)

    opts = Options(
        emphasis=not args.no_emphasis,
        color=not args.no_color,
        emphasis_class=args.emphasis_class,
        inject_style=not args.no_style,
        css=args.css,
    )

    default_ext = ".html" if args.format == "html" else ".md"
    if args.markdown:
        markdown = Path(args.markdown).read_text(encoding=args.encoding)
        out_path = Path(args.out) if args.out else Path(args.markdown).with_suffix(default_ext)
    elif args.pandoc:
        markdown = run_pandoc(args.docx, args.pandoc or args.pandoc_bin, args.to)
        out_path = Path(args.out) if args.out else Path(args.docx).with_suffix(default_ext)
    else:
        parser.error("必须提供 --markdown，或使用 --run-pandoc 由程序调用 pandoc")

    print("源 docx: %s" % args.docx)
    print("输出格式: %s | 开关: 着重号=%s 颜色=%s" % (args.format, opts.emphasis, opts.color))

    result, stats = process(
        args.docx, markdown, opts,
        output_format=args.format, pandoc_bin=args.pandoc_bin, dry_run=args.dry_run,
    )
    print("已包裹 %d 个（着重号 %d / 颜色 %d），未命中 %d 个。" % (
        stats["wrapped"], stats["by_type"]["emphasis"], stats["by_type"]["color"],
        len(stats["missed"])))

    if args.dry_run:
        return 0

    out_path.write_text(result, encoding=args.encoding)
    print("已写入: %s" % out_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
