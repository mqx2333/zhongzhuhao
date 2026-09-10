#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""restore_zhongzhuhao 的端到端测试。

流程：手工构造含 w:em / w:color 的最小 docx -> pandoc 转 markdown
      -> 运行修复 -> 断言。
另测：代码块、行内代码、链接内的短语不被误包；颜色与着重号可叠加。
"""

import subprocess
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from restore_zhongzhuhao import (
    Annotation,
    Options,
    extract_segments,
    mark_fenced_code,
    process,
    process_markdown,
    wrap_segments,
)

DOCUMENT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>这是</w:t></w:r>
      <w:r><w:rPr><w:em w:val="dot"/></w:rPr><w:t>着重</w:t></w:r>
      <w:r><w:rPr><w:em w:val="dot"/></w:rPr><w:t>强调</w:t></w:r>
      <w:r><w:t>的文字。</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>普通文字，</w:t></w:r>
      <w:r><w:rPr><w:em w:val="dot"/></w:rPr><w:t>着重</w:t></w:r>
      <w:r><w:t>也放在句子里。</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>颜色：</w:t></w:r>
      <w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>红色文字</w:t></w:r>
      <w:r><w:t>，以及</w:t></w:r>
      <w:r><w:rPr><w:em w:val="dot"/><w:color w:val="0000FF"/></w:rPr><w:t>又蓝又着重</w:t></w:r>
      <w:r><w:t>。</w:t></w:r>
    </w:p>
    <w:p><w:r><w:t>没有任何着重号的一段。</w:t></w:r></w:p>
  </w:body>
</w:document>
"""

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""


def make_docx(path: Path) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", RELS)
        zf.writestr("word/document.xml", DOCUMENT_XML)


def pandoc_to_markdown(docx_path: Path) -> str:
    proc = subprocess.run(
        ["pandoc", str(docx_path), "-t", "markdown", "--wrap=none"],
        capture_output=True, text=True, encoding="utf-8",
    )
    assert proc.returncode == 0, proc.stderr
    return proc.stdout


def test_extract_segments(tmp_path):
    docx = tmp_path / "sample.docx"
    make_docx(docx)
    segs = extract_segments(docx)
    got = [(s.text, s.annot.emphasis, s.annot.color) for s in segs]
    assert got == [
        ("着重强调", True, None),
        ("着重", True, None),
        ("红色文字", False, "#FF0000"),
        ("又蓝又着重", True, "#0000FF"),
    ], got


def test_extract_only_emphasis(tmp_path):
    docx = tmp_path / "sample.docx"
    make_docx(docx)
    segs = extract_segments(docx, want_color=False)
    assert [s.text for s in segs] == ["着重强调", "着重", "又蓝又着重"]


def test_process_markdown(tmp_path):
    docx = tmp_path / "sample.docx"
    make_docx(docx)
    md = pandoc_to_markdown(docx)

    out, stats = process_markdown(docx, md, Options())
    assert stats["wrapped"] == 4, stats
    assert stats["missed"] == [], stats
    assert stats["by_type"] == {"emphasis": 3, "color": 2}, stats

    assert '<span class="zhongzhuhao">着重强调</span>' in out
    assert '<span class="zhongzhuhao">着重</span>' in out
    assert '<span style="color:#FF0000">红色文字</span>' in out
    assert '<span class="zhongzhuhao" style="color:#0000FF">又蓝又着重</span>' in out
    assert "text-emphasis-style: dot" in out  # 样式已注入
    assert "没有任何着重号的一段。" in out


def test_no_style(tmp_path):
    docx = tmp_path / "sample.docx"
    make_docx(docx)
    md = pandoc_to_markdown(docx)
    out, _ = process_markdown(docx, md, Options(inject_style=False))
    assert "<style>" not in out


def test_process_html(tmp_path):
    docx = tmp_path / "sample.docx"
    make_docx(docx)
    md = pandoc_to_markdown(docx)

    out, stats = process(docx, md, Options(), output_format="html")
    assert stats["wrapped"] == 4, stats
    assert "<!DOCTYPE html>" in out or "<html" in out, out[:200]
    assert '<span class="zhongzhuhao">着重强调</span>' in out
    assert '<span style="color:#FF0000">红色文字</span>' in out
    assert '<span class="zhongzhuhao" style="color:#0000FF">又蓝又着重</span>' in out
    # 样式注入在 <head> 内（</head> 之前）
    low = out.lower()
    assert "text-emphasis-style: dot" in low
    assert low.find("text-emphasis-style") < low.find("</head>")


def test_skips_code_blocks_and_inline_code_and_links():
    md = (
        "先是一段有着重的文字，着重强调要加号。\n\n"
        "```text\n"
        "这里的着重强调在代码块里，不能动。\n"
        "```\n\n"
        "行内代码里 `着重强调` 也不能动。\n\n"
        "链接 [着重强调](https://example.com) 里的也不能动。\n\n"
    )
    from restore_zhongzhuhao import Segment

    seg = Segment("着重强调", Annotation(emphasis=True))
    out, stats = wrap_segments(md, [seg])
    assert stats["wrapped"] == 1, stats
    assert "<span class=\"zhongzhuhao\">着重强调</span>要加号" in out
    assert "<span class=\"zhongzhuhao\">着重强调</span>在代码块里" not in out
    assert "`着重强调`" in out
    assert "[着重强调](https://example.com)" in out
    assert "```text" in out


def test_soft_wrap_match():
    # 短语被软换行拆开时也应能匹配（空格可匹配换行）
    from restore_zhongzhuhao import Segment

    md = "一段 long 着重 强调 结束"
    seg = Segment("着重 强调", Annotation(emphasis=True))
    out, stats = wrap_segments(md, [seg])
    assert stats["wrapped"] == 1
    assert '<span class="zhongzhuhao">着重 强调</span>' in out


def test_mark_fenced_code():
    lines = ["a", "```", "code", "```", "b", "~~~", "c", "~~~", "d"]
    is_code = mark_fenced_code(lines)
    assert is_code == [False, True, True, True, False, True, True, True, False], is_code


def test_dry_run_does_not_mutate():
    from restore_zhongzhuhao import Segment

    seg = Segment("着重强调", Annotation(emphasis=True))
    out, stats = wrap_segments("内容着重强调更多", [seg], dry_run=True)
    assert stats["wrapped"] == 1
    assert "zhongzhuhao" not in out


if __name__ == "__main__":
    import tempfile

    tmp_tests = {
        test_extract_segments,
        test_extract_only_emphasis,
        test_process_markdown,
        test_no_style,
        test_process_html,
    }
    tests = [
        test_extract_segments,
        test_extract_only_emphasis,
        test_process_markdown,
        test_no_style,
        test_process_html,
        test_skips_code_blocks_and_inline_code_and_links,
        test_soft_wrap_match,
        test_mark_fenced_code,
        test_dry_run_does_not_mutate,
    ]
    for fn in tests:
        if fn in tmp_tests:
            with tempfile.TemporaryDirectory() as td:
                fn(Path(td))
        else:
            fn()
        print("ok -", fn.__name__)
    print("全部测试通过。")
