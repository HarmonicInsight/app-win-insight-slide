# -*- coding: utf-8 -*-
"""
InsightSlides EXE Build Script
PyInstallerでWindows用実行ファイルを生成
"""
import subprocess
import sys
import shutil
from pathlib import Path


# ビルド設定
APP_NAME = "InsightSlides"
ENTRY_POINT = "run_insight_slides.py"
ICON_FILE = "assets/icon.ico"  # アイコンがあれば
VERSION = "1.0.0"

# 出力ディレクトリ
DIST_DIR = Path("dist")
BUILD_DIR = Path("build")


def check_dependencies():
    """依存関係をチェック"""
    print("依存関係をチェック中...")

    required = ["pyinstaller", "python-pptx", "openpyxl"]
    missing = []

    for pkg in required:
        try:
            __import__(pkg.replace("-", "_"))
        except ImportError:
            missing.append(pkg)

    if missing:
        print(f"不足パッケージ: {', '.join(missing)}")
        print("インストール中...")
        subprocess.run([sys.executable, "-m", "pip", "install"] + missing, check=True)

    print("依存関係OK")


def clean_build():
    """ビルドディレクトリをクリーン"""
    print("クリーンアップ中...")

    for d in [DIST_DIR, BUILD_DIR]:
        if d.exists():
            shutil.rmtree(d)

    # specファイルも削除
    spec_file = Path(f"{APP_NAME}.spec")
    if spec_file.exists():
        spec_file.unlink()


def build_exe():
    """EXEをビルド"""
    print(f"\n{'='*50}")
    print(f"  {APP_NAME} v{VERSION} ビルド開始")
    print(f"{'='*50}\n")

    # PyInstallerコマンド
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name", APP_NAME,
        "--onefile",           # 単一EXE
        "--windowed",          # コンソール非表示
        "--noconfirm",         # 確認なし
        "--clean",             # クリーンビルド

        # 隠しインポート（必要に応じて追加）
        "--hidden-import", "pptx",
        "--hidden-import", "openpyxl",
        "--hidden-import", "PIL",

        # データファイル収集
        "--collect-data", "pptx",

        # エントリーポイント
        ENTRY_POINT,
    ]

    # アイコンがあれば追加
    if Path(ICON_FILE).exists():
        cmd.extend(["--icon", ICON_FILE])

    print("PyInstaller実行中...")
    print(f"コマンド: {' '.join(cmd)}\n")

    result = subprocess.run(cmd)

    if result.returncode == 0:
        exe_path = DIST_DIR / f"{APP_NAME}.exe"
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print(f"\n{'='*50}")
            print(f"  ビルド成功!")
            print(f"  出力: {exe_path}")
            print(f"  サイズ: {size_mb:.1f} MB")
            print(f"{'='*50}\n")
            return True

    print("\nビルド失敗")
    return False


def create_release_package():
    """リリースパッケージを作成"""
    print("リリースパッケージ作成中...")

    release_dir = Path(f"release/{APP_NAME}_v{VERSION}")
    release_dir.mkdir(parents=True, exist_ok=True)

    # EXEをコピー
    exe_src = DIST_DIR / f"{APP_NAME}.exe"
    if exe_src.exists():
        shutil.copy(exe_src, release_dir)

    # READMEを作成
    readme = release_dir / "README.txt"
    readme.write_text(f"""
{APP_NAME} v{VERSION}
========================

PowerPointテキストの抽出・編集・反映ツール

【使い方】
1. {APP_NAME}.exe を実行
2. PPTXファイルをドラッグ＆ドロップ
3. テキストを編集 (Excel/JSON出力も可能)
4. 保存

【動作環境】
- Windows 10/11
- .NET Framework 不要

【ライセンス】
- Free版: 3スライドまで更新可能
- Trial: 14日間全機能
- Standard/Pro: 購入ライセンスキーで解除

【サポート】
support@example.com

by Harmonic Insight
""", encoding="utf-8")

    print(f"リリースパッケージ: {release_dir}")


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description=f"{APP_NAME} EXE Builder")
    parser.add_argument("--clean", action="store_true", help="クリーンビルド")
    parser.add_argument("--release", action="store_true", help="リリースパッケージも作成")
    args = parser.parse_args()

    if args.clean:
        clean_build()

    check_dependencies()

    if build_exe():
        if args.release:
            create_release_package()
