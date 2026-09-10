# -*- coding: utf-8 -*-
"""允许 ``python -m restore_zhongzhuhao`` 运行命令行入口。"""

from .cli import main

if __name__ == "__main__":
    raise SystemExit(main())
