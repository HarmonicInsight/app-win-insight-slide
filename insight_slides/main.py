#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Insight Slides - PowerPoint編集・AI処理ツール

エントリーポイント

PyInstaller での EXE 化:
    pyinstaller insight_slides/main.py --name InsightSlides --noconsole --onefile
"""

import tkinter as tk
from .ui.main_window import MainWindow
from .config import APP_NAME


def main():
    """アプリケーションを起動"""
    root = tk.Tk()
    root.title(APP_NAME)

    # アイコン設定（存在する場合）
    try:
        from pathlib import Path
        icon_path = Path(__file__).parent / "assets" / "icon.ico"
        if icon_path.exists():
            root.iconbitmap(str(icon_path))
    except:
        pass

    # メインウィンドウを作成
    app = MainWindow(root)

    # 終了時の確認
    def on_closing():
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)

    # メインループ
    root.mainloop()


if __name__ == "__main__":
    main()
