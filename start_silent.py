# -*- coding: utf-8 -*-
"""兼容入口：桌面快捷方式指向本文件。

实际逻辑在 launcher.py（跨平台）。保留此文件名是为了让既有的
桌面快捷方式无需重建即可继续使用。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from launcher import main

if __name__ == "__main__":
    sys.exit(main())
