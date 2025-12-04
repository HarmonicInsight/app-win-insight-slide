#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
InsightSlides - メインアプリケーション起動スクリプト

PyInstaller での EXE 化時のエントリーポイント:
  pyinstaller run_app.py --name InsightSlides --noconsole
"""

import sys
from pathlib import Path

# プロジェクトルートをパスに追加（パッケージインポート用）
project_root = Path(__file__).parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from app.insightslides_app import main

if __name__ == "__main__":
    main()
