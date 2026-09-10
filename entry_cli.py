#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""CLI 打包入口（供 PyInstaller 使用）。

用法（源码运行）：
    python entry_cli.py --docx 原稿.docx --run-pandoc -o 修复后.md
"""

import sys

from restore_zhongzhuhao.cli import main

if __name__ == "__main__":
    sys.exit(main())
