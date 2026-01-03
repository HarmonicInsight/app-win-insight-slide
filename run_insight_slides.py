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

from insight_slides.main import main

if __name__ == "__main__":
    main()
