# -*- coding: utf-8 -*-
"""
Insight Slides - 設定・定数
"""

APP_NAME = "Insight Slides"
APP_VERSION = "1.0.0"

# ============== 色設定 ==============
COLORS = {
    # プライマリ
    "primary": "#2563eb",
    "primary_dark": "#1d4ed8",
    "primary_light": "#3b82f6",

    # 成功・警告・エラー
    "success": "#16a34a",
    "warning": "#ea580c",
    "error": "#dc2626",
    "info": "#0891b2",

    # 背景
    "bg": "#f8fafc",
    "bg_secondary": "#f1f5f9",
    "surface": "#ffffff",

    # テキスト
    "text": "#1e293b",
    "text_secondary": "#475569",
    "text_muted": "#64748b",

    # ボーダー
    "border": "#e2e8f0",
    "border_dark": "#cbd5e1",

    # ハイライト
    "highlight": "#fef9c3",       # 変更セルハイライト（薄黄）
    "highlight_hover": "#fef08a",

    # 差分表示
    "diff_added": "#dcfce7",      # 追加（薄緑）
    "diff_removed": "#fee2e2",    # 削除（薄赤）
    "diff_changed": "#fef3c7",    # 変更（薄オレンジ）

    # ステップインジケータ
    "step_active": "#2563eb",
    "step_inactive": "#cbd5e1",
    "step_complete": "#16a34a",
}

# ============== フォント設定 ==============
# Windows標準フォント（クリーンで読みやすい）
FONT_FAMILY = "Meiryo UI"

FONTS = {
    "title": (FONT_FAMILY, 16, "bold"),
    "heading": (FONT_FAMILY, 12, "bold"),
    "body": (FONT_FAMILY, 10),
    "small": (FONT_FAMILY, 9),
    "mono": ("Consolas", 10),
}

# ============== ステップ定義 ==============
STEPS = [
    {"id": "load", "label": "読込", "icon": "①"},
    {"id": "edit", "label": "編集・AI処理", "icon": "②"},
    {"id": "save", "label": "保存", "icon": "③"},
]

# ============== AIプリセット ==============
AI_PRESETS = {
    "翻訳（英語）": "以下のテキストを英語に翻訳してください。",
    "翻訳（日本語）": "以下のテキストを日本語に翻訳してください。",
    "要約": "以下のテキストを簡潔に要約してください。",
    "敬語変換": "以下のテキストを丁寧な敬語に変換してください。",
    "カジュアル変換": "以下のテキストをカジュアルな表現に変換してください。",
    "校正": "以下のテキストの誤字脱字を修正してください。",
    "建設用語統一": "以下のテキストを建設業界の専門用語を使用した表現に統一してください。",
}

# ============== ウィンドウサイズ ==============
WINDOW_SIZE = {
    "main": (1100, 750),
    "dialog_small": (450, 300),
    "dialog_medium": (550, 450),
    "dialog_large": (800, 600),
}
