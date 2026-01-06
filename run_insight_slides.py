#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Insight Slides - 起動スクリプト

Usage:
    python run_insight_slides.py

PyInstaller での EXE 化:
    pyinstaller run_insight_slides.py --name InsightSlides --noconsole --onefile
"""

import sys
from pathlib import Path

# プロジェクトルートをパスに追加
project_root = Path(__file__).parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from core.common_bootstrap import add_insight_common_paths

common_dir = add_insight_common_paths()
if common_dir:
    try:
        import __init__ as i18n
        print("[common] i18n loaded:", hasattr(i18n, "t"))
    except Exception as e:
        print("[common] i18n load failed:", e)
else:
    print("[common] insight-common not found")

from insight_slides.main import main

if __name__ == "__main__":
    main()
