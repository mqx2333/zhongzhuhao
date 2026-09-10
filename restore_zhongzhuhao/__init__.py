# -*- coding: utf-8 -*-
"""restore-zhongzhuhao：把 docx 中被 pandoc 忽略的着重号与字体颜色补回 Markdown / HTML。"""

from .core import (
    Annotation,
    DEFAULT_EMPHASIS_CLASS,
    Options,
    Segment,
    default_css,
    extract_segments,
    inject_html_style,
    mark_fenced_code,
    markdown_to_html,
    process,
    process_markdown,
    run_pandoc,
    wrap_segments,
)

__version__ = "0.5.0"

__all__ = [
    "Annotation",
    "DEFAULT_EMPHASIS_CLASS",
    "Options",
    "Segment",
    "default_css",
    "extract_segments",
    "inject_html_style",
    "mark_fenced_code",
    "markdown_to_html",
    "process",
    "process_markdown",
    "run_pandoc",
    "wrap_segments",
    "__version__",
]
