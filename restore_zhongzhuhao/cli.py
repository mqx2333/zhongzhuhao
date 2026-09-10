# -*- coding: utf-8 -*-
"""命令行入口：``python -m restore_zhongzhuhao`` 或安装后的 ``restore-zhongzhuhao``。"""

from __future__ import annotations

import argparse
from pathlib import Path

from .core import DEFAULT_EMPHASIS_CLASS, Options, process, run_pandoc


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="restore-zhongzhuhao",
        description="辅助 pandoc：把 docx 中被忽略的着重号与字体颜色以内联 HTML 补回 Markdown / HTML。",
    )
    parser.add_argument("--docx", required=True, help="源 Word .docx 文件")
    parser.add_argument("--markdown", help="pandoc 生成的 Markdown 文件")
    parser.add_argument("-o", "--out", help="输出文件；省略则写回 --markdown / 同名")
    parser.add_argument("--format", choices=["markdown", "html"], default="markdown",
                        help="输出格式（默认 markdown；html 会用 pandoc 转成独立网页）")
    parser.add_argument("--run-pandoc", dest="pandoc", nargs="?", const="pandoc",
                        metavar="PANDOC", help="先调用 pandoc 把 docx 转成 markdown 再修复")
    parser.add_argument("--pandoc-bin", default="pandoc", help="pandoc 可执行文件（HTML 转换用）")
    parser.add_argument("--to", default="markdown", help="pandoc 输出格式（默认 markdown）")
    parser.add_argument("--no-emphasis", action="store_true", help="不处理着重号")
    parser.add_argument("--no-color", action="store_true", help="不处理字体颜色")
    parser.add_argument("--emphasis-class", default=DEFAULT_EMPHASIS_CLASS,
                        help="着重号 span 的 class 名（默认 %s）" % DEFAULT_EMPHASIS_CLASS)
    parser.add_argument("--css", default=None, help="自定义 .zhongzhuhao 样式")
    parser.add_argument("--no-style", action="store_true", help="不注入样式块")
    parser.add_argument("--dry-run", action="store_true", help="只打印将处理的短语，不写文件")
    parser.add_argument("--encoding", default="utf-8", help="文件读写编码（默认 utf-8）")
    return parser


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)

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
        raise SystemExit("必须提供 --markdown，或使用 --run-pandoc 由程序调用 pandoc")

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
    raise SystemExit(main())
