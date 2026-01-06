# core/common_bootstrap.py
from __future__ import annotations

import sys
from pathlib import Path
from typing import Optional

def _repo_root() -> Path:
    """
    Repo root を推定する。
    - 開発中: このファイルの位置から辿る
    - PyInstaller: sys._MEIPASS や exe の場所を起点にする
    """
    # PyInstaller onefile/onedir対応
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        return Path(meipass).resolve()

    # 通常実行: core/ から1つ上をルートとみなす
    return Path(__file__).resolve().parent.parent

def add_insight_common_paths() -> Optional[Path]:
    """
    insight-common を見つけて必要なサブディレクトリを sys.path に追加する。
    戻り値: 見つかった insight-common のパス（見つからなければ None）
    """
    root = _repo_root()

    # 開発中: repo直下の insight-common
    cand = root / "insight-common"
    # PyInstaller: datas で "insight-common" を同梱した場合
    if not cand.exists():
        cand = root / "insight-common"

    if not cand.exists():
        return None

    # InsightPy方式：各モジュールの __init__.py を直接 import する前提
    paths = [
        cand / "i18n",
        cand / "errors",
        cand / "utils" / "python",
        cand / "license" / "python",
    ]
    for p in paths:
        if p.exists():
            sys.path.insert(0, str(p.resolve()))

    return cand
