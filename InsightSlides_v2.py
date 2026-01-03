# -*- coding: utf-8 -*-
"""
Insight Slides v2.0 - PowerPoint Text Extract & Update Tool
統合版: 旧UI + グリッド編集 + 比較機能 + フィルタ

by Harmonic Insight

特徴:
- 抽出/更新モード切替
- インライングリッド編集
- PPTX比較機能
- フィルタ機能
- 統一ライセンス形式 (INS-SLIDE-{TIER}-XXXX-XXXX-CC)
- 折りたたみ可能なオプション
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, simpledialog
import pptx
import openpyxl
from openpyxl.styles import Font as XLFont, PatternFill
import os
import re
import csv
import json
import hashlib
import random
import string
import webbrowser
import traceback
from datetime import datetime, timedelta
import threading
from pathlib import Path
from typing import Dict, Tuple, List, Optional
import shutil

# ============== App Info ==============
APP_VERSION = "2.0.0"
APP_NAME = "Insight Slides"

# ============== Config Paths ==============
CONFIG_DIR = Path.home() / ".insightslides"
CONFIG_FILE = CONFIG_DIR / "config.json"
LICENSE_FILE = CONFIG_DIR / "license.key"
ERROR_LOG_FILE = CONFIG_DIR / "error_log.txt"

# ============== Support Links ==============
SUPPORT_LINKS = {
    "faq": "https://example.com/insightslides/faq",
    "tutorial": "https://example.com/insightslides/tutorial",
    "purchase": "https://example.com/insightslides/purchase",
    "contact": "mailto:support@example.com",
}

# ============== Internationalization (i18n) ==============
LANGUAGES = {
    'en': {
        'app_subtitle': 'Extract → Edit → Update PowerPoint Text',
        'welcome_title': 'Welcome to Insight Slides!',
        'mode_extract': 'Extract Mode',
        'mode_update': 'Update Mode',
        'mode_extract_short': 'Extract Text',
        'mode_update_short': 'Overwrite',
        'panel_mode': 'Mode Selection',
        'panel_file': 'File Operations',
        'panel_settings': 'Settings',
        'panel_status': 'Status',
        'panel_output': 'Extracted Data',
        'panel_extract_options': 'Extract Options',
        'panel_update_options': 'Update Options',
        'panel_extract_run': 'Run Extract',
        'panel_update_run': 'Run Update',
        'panel_pro_features': 'Pro Features',
        'btn_single_file': 'Select File',
        'btn_batch_folder': 'Folder Batch',
        'btn_from_excel': 'From Excel',
        'btn_from_json': 'From JSON',
        'btn_batch_update': 'Folder Batch Update',
        'btn_diff_preview': 'Diff Preview',
        'btn_compare_pptx': 'Compare PPTX',
        'btn_cancel': 'Stop',
        'btn_clear': 'Clear Log',
        'btn_copy': 'Copy Log',
        'btn_license': 'License',
        'btn_activate': 'Activate',
        'btn_deactivate': 'Deactivate',
        'btn_purchase': 'Purchase',
        'btn_close': 'Close',
        'btn_start': 'Get Started',
        'btn_filter': 'Filter',
        'btn_clear_filter': 'Clear Filter',
        'setting_output_format': 'Output Format:',
        'setting_include_meta': 'Include file name & date',
        'setting_auto_backup': 'Auto backup before update',
        'chk_include_notes': 'Include Speaker Notes',
        'format_tab': 'Tab-separated',
        'format_csv': 'CSV',
        'format_excel': 'Excel',
        'status_waiting': 'Waiting...',
        'status_processing': 'Processing...',
        'status_complete': 'Complete',
        'status_cancelled': 'Cancelled',
        'status_error': 'Error',
        'msg_extract_desc': 'Extract text from PowerPoint files.',
        'msg_update_desc': 'Apply edited text back to PowerPoint.',
        'msg_update_limit': 'Update: First {0} slides only\nUpgrade to Standard for unlimited!',
        'msg_processing_file': 'Processing: {0}',
        'msg_saved': 'Saved: {0}',
        'msg_extracted': 'Extracted: {0} items from {1} slides',
        'msg_updated': 'Updated: {0} items, Skipped: {1}',
        'msg_no_pptx': 'No PPTX files found',
        'msg_no_data': 'No update data found',
        'msg_copied': 'Copied to clipboard',
        'license_title': 'License Management',
        'license_current': 'Current License',
        'license_enter_key': 'Enter License Key:',
        'license_activated': '{0} has been activated',
        'license_deactivated': 'License deactivated',
        'license_invalid': 'Invalid license key',
        'license_enter_prompt': 'Please enter a license key',
        'upgrade_title': 'Upgrade',
        'dialog_confirm': 'Confirm',
        'dialog_error': 'Error',
        'dialog_complete': 'Complete',
        'header_slide': 'Slide',
        'header_id': 'Object ID',
        'header_type': 'Type',
        'header_text': 'Text Content',
        'header_filename': 'Filename',
        'header_datetime': 'Extracted At',
        'diff_title': 'Diff Preview',
        'menu_help': 'Help',
        'menu_guide': 'User Guide',
        'menu_faq': 'FAQ',
        'menu_license': 'License Management',
        'menu_about': 'About',
        'lang_menu': 'Language',
        'font_size_menu': 'Font Size',
        'font_size_small': 'Small',
        'font_size_medium': 'Medium',
        'font_size_large': 'Large',
        'advanced_options': 'Advanced Options',
        'type_notes': 'Notes',
        'filter_placeholder': 'Filter text...',
    },
    'ja': {
        'app_subtitle': 'PowerPointテキストを抽出 → 編集 → 反映',
        'welcome_title': 'Insight Slides へようこそ！',
        'mode_extract': '抽出モード',
        'mode_update': '更新モード',
        'mode_extract_short': 'テキスト抽出',
        'mode_update_short': '上書き更新',
        'panel_mode': 'モード選択',
        'panel_file': 'ファイル操作',
        'panel_settings': '処理設定',
        'panel_status': '処理状況',
        'panel_output': '抽出結果',
        'panel_extract_options': '抽出オプション',
        'panel_update_options': '更新オプション',
        'panel_extract_run': '抽出実行',
        'panel_update_run': '更新実行',
        'panel_pro_features': '拡張機能',
        'btn_single_file': 'ファイル選択',
        'btn_batch_folder': 'フォルダー一括',
        'btn_from_excel': 'Excelから更新',
        'btn_from_json': 'JSONから更新',
        'btn_batch_update': 'フォルダー一括更新',
        'btn_diff_preview': '差分プレビュー',
        'btn_compare_pptx': 'PPTX比較',
        'btn_cancel': '中止',
        'btn_clear': 'ログクリア',
        'btn_copy': 'ログコピー',
        'btn_license': 'ライセンス',
        'btn_activate': 'アクティベート',
        'btn_deactivate': 'ライセンス解除',
        'btn_purchase': '購入ページ',
        'btn_close': '閉じる',
        'btn_start': '始める',
        'btn_filter': 'フィルタ',
        'btn_clear_filter': 'クリア',
        'setting_output_format': '出力形式:',
        'setting_include_meta': 'ファイル名・日時を含める',
        'setting_auto_backup': '更新前に自動バックアップ',
        'chk_include_notes': 'スピーカーノート含む',
        'format_tab': 'タブ区切り',
        'format_csv': 'CSV形式',
        'format_excel': 'Excel形式',
        'status_waiting': '処理待機中...',
        'status_processing': '処理中...',
        'status_complete': '完了',
        'status_cancelled': 'キャンセルされました',
        'status_error': 'エラー',
        'msg_extract_desc': 'PowerPointからテキストを抽出します。',
        'msg_update_desc': '編集したファイルの変更をPowerPointに反映します。',
        'msg_update_limit': '更新機能: 最初の{0}スライドのみ\nStandard版で無制限に！',
        'msg_processing_file': '処理中: {0}',
        'msg_saved': '保存完了: {0}',
        'msg_extracted': '抽出: {0}件 / スライド: {1}枚',
        'msg_updated': '更新: {0}件 / スキップ: {1}件',
        'msg_no_pptx': 'PPTXファイルが見つかりません',
        'msg_no_data': '更新データがありません',
        'msg_copied': 'クリップボードにコピーしました',
        'license_title': 'ライセンス管理',
        'license_current': '現在のライセンス',
        'license_enter_key': 'ライセンスキー:',
        'license_activated': '{0}版がアクティベートされました',
        'license_deactivated': 'ライセンスを解除しました',
        'license_invalid': '無効なライセンスキーです',
        'license_enter_prompt': 'ライセンスキーを入力してください',
        'upgrade_title': 'アップグレード',
        'dialog_confirm': '確認',
        'dialog_error': 'エラー',
        'dialog_complete': '完了',
        'header_slide': 'スライド番号',
        'header_id': 'オブジェクトID',
        'header_type': 'タイプ',
        'header_text': 'テキスト内容',
        'header_filename': 'ファイル名',
        'header_datetime': '抽出日時',
        'diff_title': '差分プレビュー',
        'menu_help': 'ヘルプ',
        'menu_guide': '使い方ガイド',
        'menu_faq': 'よくある質問',
        'menu_license': 'ライセンス管理',
        'menu_about': 'バージョン情報',
        'lang_menu': '言語 / Language',
        'font_size_menu': '文字サイズ',
        'font_size_small': '小',
        'font_size_medium': '中',
        'font_size_large': '大',
        'advanced_options': '詳細オプション',
        'type_notes': 'ノート',
        'filter_placeholder': 'フィルタ...',
    },
}

_current_lang = 'ja'

def t(key: str, *args) -> str:
    text = LANGUAGES.get(_current_lang, LANGUAGES['ja']).get(key, key)
    if args:
        return text.format(*args)
    return text

def set_language(lang: str):
    global _current_lang
    if lang in LANGUAGES:
        _current_lang = lang

def get_language() -> str:
    return _current_lang


# ============== ライセンス設定（統一形式） ==============
LICENSE_SECRET = "HarmonicInsight2025"
PRODUCT_CODE = "SLIDE"

class LicenseTier:
    FREE = "FREE"
    TRIAL = "TRIAL"
    STD = "STD"
    PRO = "PRO"
    ENT = "ENT"

TIERS = {
    LicenseTier.FREE: {'name': 'Free', 'name_ja': '無料版', 'badge': 'Free', 'update_limit': 3, 'batch': False, 'pro': False},
    LicenseTier.TRIAL: {'name': 'Trial', 'name_ja': 'トライアル', 'badge': 'Trial', 'update_limit': None, 'batch': True, 'pro': True, 'days': 14},
    LicenseTier.STD: {'name': 'Standard', 'name_ja': 'スタンダード', 'badge': '📘 Standard', 'update_limit': None, 'batch': True, 'pro': False},
    LicenseTier.PRO: {'name': 'Professional', 'name_ja': 'プロフェッショナル', 'badge': '⭐ Pro', 'update_limit': None, 'batch': True, 'pro': True},
    LicenseTier.ENT: {'name': 'Enterprise', 'name_ja': 'エンタープライズ', 'badge': '🏢 Enterprise', 'update_limit': None, 'batch': True, 'pro': True},
}


def _generate_checksum(key_body: str) -> str:
    return hashlib.sha256(f"{key_body}{LICENSE_SECRET}".encode()).hexdigest()[:2].upper()


def validate_license_key(license_key: str) -> Tuple[bool, str, Optional[str]]:
    """
    統一形式でライセンスキーを検証
    形式: INS-SLIDE-{TIER}-XXXX-XXXX-CC
    Returns: (is_valid, tier, expires)
    """
    if not license_key:
        return False, LicenseTier.FREE, None

    key = license_key.strip().upper()
    parts = key.split("-")

    # 形式チェック: INS-SLIDE-TIER-XXXX-XXXX-CC (6パーツ)
    if len(parts) != 6:
        return False, LicenseTier.FREE, None

    prefix, product, tier_str, part1, part2, checksum = parts

    if prefix != "INS" or product != PRODUCT_CODE:
        return False, LicenseTier.FREE, None

    if tier_str not in [LicenseTier.FREE, LicenseTier.TRIAL, LicenseTier.STD, LicenseTier.PRO, LicenseTier.ENT]:
        return False, LicenseTier.FREE, None

    # チェックサム検証
    key_body = f"{prefix}-{product}-{tier_str}-{part1}-{part2}"
    expected_checksum = _generate_checksum(key_body)
    if checksum != expected_checksum:
        return False, LicenseTier.FREE, None

    # 有効期限計算
    expires = None
    tier_config = TIERS.get(tier_str, TIERS[LicenseTier.FREE])
    if tier_config.get('days'):
        expires = (datetime.now() + timedelta(days=tier_config['days'])).strftime("%Y-%m-%d")

    return True, tier_str, expires


def generate_license_key(tier: str) -> str:
    """ライセンスキー生成: INS-SLIDE-{TIER}-XXXX-XXXX-CC"""
    chars = string.ascii_uppercase + string.digits
    part1 = ''.join(random.choices(chars, k=4))
    part2 = ''.join(random.choices(chars, k=4))
    key_body = f"INS-{PRODUCT_CODE}-{tier}-{part1}-{part2}"
    checksum = _generate_checksum(key_body)
    return f"{key_body}-{checksum}"


class LicenseManager:
    def __init__(self):
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        self.license_info = self._load_license()

    def _load_license(self) -> Dict:
        if LICENSE_FILE.exists():
            try:
                with open(LICENSE_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    if data.get('key'):
                        is_valid, tier, expires = validate_license_key(data['key'])
                        if is_valid:
                            return {'type': tier, 'key': data['key'], 'expires': expires}
            except:
                pass
        return {'type': LicenseTier.FREE, 'key': '', 'expires': None}

    def _save_license(self, data: Dict):
        with open(LICENSE_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def activate(self, key: str) -> Tuple[bool, str]:
        if not key:
            return False, t('license_enter_prompt')

        is_valid, tier, expires = validate_license_key(key.strip())
        if not is_valid:
            return False, t('license_invalid')

        self.license_info = {'type': tier, 'key': key.strip().upper(), 'expires': expires}
        self._save_license(self.license_info)

        tier_info = TIERS.get(tier, TIERS[LicenseTier.FREE])
        name = tier_info['name_ja'] if get_language() == 'ja' else tier_info['name']
        return True, t('license_activated', name)

    def deactivate(self):
        self.license_info = {'type': LicenseTier.FREE, 'key': '', 'expires': None}
        if LICENSE_FILE.exists():
            LICENSE_FILE.unlink()

    def get_tier(self) -> str:
        return self.license_info.get('type', LicenseTier.FREE)

    def get_tier_info(self) -> Dict:
        return TIERS.get(self.get_tier(), TIERS[LicenseTier.FREE])

    def get_update_limit(self) -> Optional[int]:
        return self.get_tier_info().get('update_limit')

    def can_batch(self) -> bool:
        return self.get_tier_info().get('batch', False)

    def is_pro(self) -> bool:
        return self.get_tier_info().get('pro', False)


# ============== デザインシステム ==============
COLOR_PALETTE = {
    "bg_primary": "#FAFBFC", "bg_secondary": "#F3F4F6", "bg_elevated": "#FFFFFF",
    "text_primary": "#1A202C", "text_secondary": "#4A5568", "text_muted": "#718096",
    "brand_primary": "#1E40AF", "brand_update": "#0D9488",
    "success": "#059669", "warning": "#D97706", "error": "#DC2626",
    "border_light": "#E5E7EB", "border_medium": "#D1D5DB",
    "diff_changed": "#FEF3C7", "diff_added": "#D1FAE5", "diff_removed": "#FEE2E2",
}

FONT_FAMILY = "Yu Gothic UI"

def get_fonts(size_preset: str = 'medium') -> dict:
    base = {'small': 10, 'medium': 11, 'large': 13}.get(size_preset, 11)
    return {
        "display": (FONT_FAMILY, base + 9, "bold"),
        "heading": (FONT_FAMILY, base + 3, "bold"),
        "subheading": (FONT_FAMILY, base + 1, "bold"),
        "body": (FONT_FAMILY, base, "normal"),
        "body_bold": (FONT_FAMILY, base, "bold"),
        "caption": (FONT_FAMILY, base - 1, "normal"),
        "small": (FONT_FAMILY, base - 2, "normal"),
    }

FONTS = get_fonts('medium')
SPACING = {"xs": 2, "sm": 4, "md": 8, "lg": 12, "xl": 16}


class ConfigManager:
    DEFAULT = {
        'language': 'ja', 'output_format': 'excel', 'include_metadata': True,
        'auto_backup': True, 'last_directory': '', 'font_size': 'medium',
        'advanced_expanded': False,
    }

    def __init__(self):
        global FONTS
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        self.config = self._load()
        set_language(self.config.get('language', 'ja'))
        FONTS = get_fonts(self.config.get('font_size', 'medium'))

    def _load(self) -> Dict:
        if CONFIG_FILE.exists():
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    return {**self.DEFAULT, **json.load(f)}
            except:
                pass
        return self.DEFAULT.copy()

    def save(self):
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, ensure_ascii=False, indent=2)

    def get(self, key: str, default=None):
        return self.config.get(key, default)

    def set(self, key: str, value):
        self.config[key] = value
        self.save()


def save_error_log(error: Exception, context: str = ""):
    try:
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        with open(ERROR_LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(f"\n{'='*60}\n{datetime.now()}\n{context}\n{error}\n{traceback.format_exc()}\n")
    except:
        pass


# ============== グリッドUI (Undo/Redo対応) ==============
class UndoManager:
    def __init__(self, max_history: int = 50):
        self.undo_stack: List[Dict] = []
        self.redo_stack: List[Dict] = []
        self.max_history = max_history

    def push(self, action: Dict):
        self.undo_stack.append(action)
        self.redo_stack.clear()
        if len(self.undo_stack) > self.max_history:
            self.undo_stack.pop(0)

    def undo(self) -> Optional[Dict]:
        if not self.undo_stack:
            return None
        action = self.undo_stack.pop()
        self.redo_stack.append(action)
        return action

    def redo(self) -> Optional[Dict]:
        if not self.redo_stack:
            return None
        action = self.redo_stack.pop()
        self.undo_stack.append(action)
        return action

    def clear(self):
        self.undo_stack.clear()
        self.redo_stack.clear()


class EditableGrid(ttk.Frame):
    """インライン編集対応グリッド + フィルタ機能"""

    def __init__(self, parent, on_change=None, **kwargs):
        super().__init__(parent, **kwargs)

        self.on_change = on_change
        self.undo_manager = UndoManager()
        self._edit_widget = None
        self._editing_item = None
        self._editing_column = None
        self._all_data: List[Dict] = []
        self._filter_text = ""

        self._create_widgets()
        self._setup_bindings()

    def _create_widgets(self):
        # ツールバー
        toolbar = ttk.Frame(self)
        toolbar.pack(fill="x", pady=(0, 5))

        # フィルタ
        ttk.Label(toolbar, text="🔍", font=FONTS["body"]).pack(side="left")
        self.filter_var = tk.StringVar()
        self.filter_entry = ttk.Entry(toolbar, textvariable=self.filter_var, width=20)
        self.filter_entry.pack(side="left", padx=(5, 5))
        self.filter_var.trace_add("write", lambda *args: self._apply_filter())

        ttk.Button(toolbar, text=t('btn_clear_filter'), command=self._clear_filter, width=6).pack(side="left")

        # スペーサー
        ttk.Frame(toolbar).pack(side="left", fill="x", expand=True)

        # 一括置換ボタン
        ttk.Button(toolbar, text="🔄 一括置換", command=self._show_replace_dialog, width=10).pack(side="left", padx=2)

        # Undo/Redo
        self.undo_btn = ttk.Button(toolbar, text="↩ 元に戻す", command=self._do_undo, width=10)
        self.undo_btn.pack(side="left", padx=2)
        self.redo_btn = ttk.Button(toolbar, text="↪ やり直し", command=self._do_redo, width=10)
        self.redo_btn.pack(side="left", padx=2)

        # Treeview
        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill="both", expand=True)

        columns = ("slide", "id", "type", "text")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings")

        self.tree.heading("slide", text=t('header_slide'))
        self.tree.heading("id", text=t('header_id'))
        self.tree.heading("type", text=t('header_type'))
        self.tree.heading("text", text=t('header_text'))

        self.tree.column("slide", width=80, anchor="center")
        self.tree.column("id", width=100)
        self.tree.column("type", width=100)
        self.tree.column("text", width=500)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        self.tree.tag_configure("modified", background=COLOR_PALETTE["diff_changed"])
        self.tree.tag_configure("filtered", background=COLOR_PALETTE["diff_added"])

    def _setup_bindings(self):
        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<Control-z>", lambda e: self._do_undo())
        self.tree.bind("<Control-y>", lambda e: self._do_redo())

    def _on_double_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return

        column = self.tree.identify_column(event.x)
        item = self.tree.identify_row(event.y)

        if not item or column != "#4":  # textカラムのみ編集可能
            return

        self._start_edit(item, "text")

    def _start_edit(self, item: str, column: str):
        bbox = self.tree.bbox(item, column)
        if not bbox:
            return

        current_value = self.tree.set(item, column)
        self._editing_item = item
        self._editing_column = column

        self._edit_widget = tk.Entry(self.tree, font=FONTS["body"])
        self._edit_widget.insert(0, current_value)
        self._edit_widget.select_range(0, tk.END)
        self._edit_widget.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        self._edit_widget.focus_set()

        self._edit_widget.bind("<Return>", lambda e: self._finish_edit())
        self._edit_widget.bind("<Escape>", lambda e: self._cancel_edit())
        self._edit_widget.bind("<FocusOut>", lambda e: self._finish_edit())

    def _finish_edit(self):
        if not self._edit_widget or not self._editing_item:
            return

        new_value = self._edit_widget.get()
        old_value = self.tree.set(self._editing_item, self._editing_column)

        if new_value != old_value:
            self.tree.set(self._editing_item, self._editing_column, new_value)
            self.tree.item(self._editing_item, tags=("modified",))

            self.undo_manager.push({
                "type": "edit", "item": self._editing_item,
                "column": self._editing_column, "old": old_value, "new": new_value
            })

            # 元データも更新
            idx = self.tree.index(self._editing_item)
            if idx < len(self._all_data):
                self._all_data[idx]["text"] = new_value

            if self.on_change:
                self.on_change(self._editing_item, self._editing_column, new_value)

        self._cancel_edit()

    def _cancel_edit(self):
        if self._edit_widget:
            self._edit_widget.destroy()
            self._edit_widget = None
        self._editing_item = None
        self._editing_column = None

    def _do_undo(self):
        action = self.undo_manager.undo()
        if action:
            self.tree.set(action["item"], action["column"], action["old"])

    def _do_redo(self):
        action = self.undo_manager.redo()
        if action:
            self.tree.set(action["item"], action["column"], action["new"])
            self.tree.item(action["item"], tags=("modified",))

    def _apply_filter(self):
        self._filter_text = self.filter_var.get().lower()
        self._refresh_display()

    def _clear_filter(self):
        self.filter_var.set("")
        self._filter_text = ""
        self._refresh_display()

    def _refresh_display(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for row in self._all_data:
            if self._filter_text:
                if self._filter_text not in str(row.get("text", "")).lower():
                    continue

            self.tree.insert("", "end", values=(
                row.get("slide", ""),
                row.get("id", ""),
                row.get("type", ""),
                row.get("text", "")
            ))

    def _show_replace_dialog(self):
        dialog = tk.Toplevel(self)
        dialog.title("一括置換")
        dialog.geometry("400x150")
        dialog.transient(self)
        dialog.grab_set()

        ttk.Label(dialog, text="検索:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        find_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=find_var, width=40).grid(row=0, column=1, padx=10, pady=10)

        ttk.Label(dialog, text="置換:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        replace_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=replace_var, width=40).grid(row=1, column=1, padx=10, pady=10)

        def do_replace():
            find_text = find_var.get()
            replace_text = replace_var.get()
            if not find_text:
                return

            count = 0
            for item in self.tree.get_children():
                old_text = self.tree.set(item, "text")
                if find_text in old_text:
                    new_text = old_text.replace(find_text, replace_text)
                    self.tree.set(item, "text", new_text)
                    self.tree.item(item, tags=("modified",))
                    count += 1

            dialog.destroy()
            messagebox.showinfo("完了", f"{count} 件を置換しました")

        ttk.Button(dialog, text="置換", command=do_replace).grid(row=2, column=1, pady=10, sticky="e")

    def load_data(self, data: List[Dict]):
        self._all_data = data.copy()
        self.undo_manager.clear()
        self._refresh_display()

    def clear(self):
        self._all_data = []
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.undo_manager.clear()

    def get_data(self) -> List[Dict]:
        result = []
        for item in self.tree.get_children():
            result.append({
                "slide": self.tree.set(item, "slide"),
                "id": self.tree.set(item, "id"),
                "type": self.tree.set(item, "type"),
                "text": self.tree.set(item, "text"),
            })
        return result


# ============== 比較機能 ==============
class CompareDialog:
    def __init__(self, parent, callback):
        self.callback = callback
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("PPTX比較")
        self.dialog.geometry("600x280")
        self.dialog.transient(parent)
        self.dialog.grab_set()

        self._create_widgets()

    def _create_widgets(self):
        frame = ttk.Frame(self.dialog, padding=20)
        frame.pack(fill='both', expand=True)

        ttk.Label(frame, text="2つのPowerPointファイルを比較", font=FONTS["heading"]).pack(anchor='w', pady=(0, 15))

        # ファイル1
        f1 = ttk.Frame(frame)
        f1.pack(fill='x', pady=5)
        ttk.Label(f1, text="元ファイル:", width=12).pack(side='left')
        self.file1_var = tk.StringVar()
        ttk.Entry(f1, textvariable=self.file1_var, width=45).pack(side='left', padx=5)
        ttk.Button(f1, text="参照", command=lambda: self._browse(self.file1_var)).pack(side='left')

        # ファイル2
        f2 = ttk.Frame(frame)
        f2.pack(fill='x', pady=5)
        ttk.Label(f2, text="新ファイル:", width=12).pack(side='left')
        self.file2_var = tk.StringVar()
        ttk.Entry(f2, textvariable=self.file2_var, width=45).pack(side='left', padx=5)
        ttk.Button(f2, text="参照", command=lambda: self._browse(self.file2_var)).pack(side='left')

        # オプション
        opt = ttk.Frame(frame)
        opt.pack(fill='x', pady=15)
        self.ignore_ws = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt, text="空白の違いを無視", variable=self.ignore_ws).pack(side='left')

        # ボタン
        btn = ttk.Frame(frame)
        btn.pack(fill='x', pady=10)
        ttk.Button(btn, text="キャンセル", command=self.dialog.destroy).pack(side='left')
        tk.Button(btn, text="比較実行", bg=COLOR_PALETTE["brand_primary"], fg="white",
                  command=self._execute).pack(side='left', padx=10)

    def _browse(self, var):
        path = filedialog.askopenfilename(filetypes=[("PowerPoint", "*.pptx")])
        if path:
            var.set(path)

    def _execute(self):
        f1, f2 = self.file1_var.get(), self.file2_var.get()
        if not f1 or not f2:
            messagebox.showwarning("警告", "2つのファイルを選択してください")
            return
        self.callback(f1, f2, self.ignore_ws.get())
        self.dialog.destroy()


class CompareResultWindow:
    def __init__(self, parent, file1_name, file2_name, diff_data, stats, on_apply=None):
        self.window = tk.Toplevel(parent)
        self.window.title(f"比較結果: {file1_name} ↔ {file2_name}")
        self.window.geometry("1100x700")

        self.diff_data = diff_data
        self.on_apply = on_apply
        self.selections = {}

        for i, row in enumerate(diff_data):
            if row["status"] == "変更":
                self.selections[i] = None
            elif row["status"] == "追加":
                self.selections[i] = "after"
            elif row["status"] == "削除":
                self.selections[i] = "before"
            else:
                self.selections[i] = "same"

        self._create_widgets(stats, file1_name, file2_name)

    def _create_widgets(self, stats, f1, f2):
        # 統計
        top = ttk.Frame(self.window, padding=10)
        top.pack(fill='x')
        ttk.Label(top, text=f"📊 一致 {stats['same']} | 変更 {stats['changed']} | 追加 {stats['added']} | 削除 {stats['removed']}",
                  font=FONTS["heading"]).pack(side='left')

        ttk.Button(top, text="CSVエクスポート", command=self._export_csv).pack(side='right')

        # グリッド
        grid_frame = ttk.Frame(self.window, padding=10)
        grid_frame.pack(fill='both', expand=True)

        cols = ("select", "slide", "id", "status", "before", "after")
        self.tree = ttk.Treeview(grid_frame, columns=cols, show="headings")

        self.tree.heading("select", text="採用")
        self.tree.heading("slide", text="スライド")
        self.tree.heading("id", text="ID")
        self.tree.heading("status", text="状態")
        self.tree.heading("before", text=f"元: {f1}")
        self.tree.heading("after", text=f"新: {f2}")

        self.tree.column("select", width=60, anchor="center")
        self.tree.column("slide", width=60, anchor="center")
        self.tree.column("id", width=80)
        self.tree.column("status", width=60, anchor="center")
        self.tree.column("before", width=350)
        self.tree.column("after", width=350)

        vsb = ttk.Scrollbar(grid_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y')
        self.tree.pack(fill='both', expand=True)

        self.tree.tag_configure("same", background=COLOR_PALETTE["bg_secondary"])
        self.tree.tag_configure("changed", background=COLOR_PALETTE["diff_changed"])
        self.tree.tag_configure("added", background=COLOR_PALETTE["diff_added"])
        self.tree.tag_configure("removed", background=COLOR_PALETTE["diff_removed"])

        self.tree.bind("<Button-1>", self._on_click)
        self.item_map = {}
        self._refresh()

        # ボタン
        bottom = ttk.Frame(self.window, padding=10)
        bottom.pack(fill='x')
        ttk.Button(bottom, text="全て元", command=lambda: self._select_all("before")).pack(side='left', padx=2)
        ttk.Button(bottom, text="全て新", command=lambda: self._select_all("after")).pack(side='left', padx=2)
        tk.Button(bottom, text="選択を反映 →", bg=COLOR_PALETTE["brand_primary"], fg="white",
                  command=self._apply).pack(side='right', padx=5)
        ttk.Button(bottom, text="閉じる", command=self.window.destroy).pack(side='right')

    def _refresh(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.item_map = {}

        for i, row in enumerate(self.diff_data):
            sel = self.selections.get(i)
            sel_text = {"before": "◀ 元", "after": "新 ▶", "same": "─"}.get(sel, "")
            tag = {"一致": "same", "変更": "changed", "追加": "added", "削除": "removed"}.get(row["status"], "same")

            before = (row.get("before") or "").replace("\n", " ↵ ")[:50]
            after = (row.get("after") or "").replace("\n", " ↵ ")[:50]

            item_id = self.tree.insert("", "end", values=(
                sel_text, row["slide"], row.get("id", ""), row["status"], before, after
            ), tags=(tag,))
            self.item_map[item_id] = i

    def _on_click(self, event):
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not item:
            return

        idx = self.item_map.get(item)
        if idx is None or self.diff_data[idx]["status"] == "一致":
            return

        current = self.selections.get(idx)
        if col == "#5":
            self.selections[idx] = "before"
        elif col == "#6":
            self.selections[idx] = "after"
        else:
            self.selections[idx] = "after" if current == "before" else "before"
        self._refresh()

    def _select_all(self, choice):
        for i, row in enumerate(self.diff_data):
            if row["status"] != "一致":
                self.selections[i] = choice
        self._refresh()

    def _apply(self):
        selected = []
        for i, row in enumerate(self.diff_data):
            sel = self.selections.get(i)
            if sel in ("before", "after"):
                text = row["before"] if sel == "before" else row["after"]
                selected.append({"slide": row["slide"], "id": row.get("id"), "text": text})

        if not selected:
            messagebox.showwarning("警告", "反映する項目がありません")
            return

        if self.on_apply:
            self.on_apply(selected)
            messagebox.showinfo("完了", f"{len(selected)} 件を反映しました")
            self.window.destroy()

    def _export_csv(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if not path:
            return

        with open(path, 'w', newline='', encoding='utf-8-sig') as f:
            w = csv.writer(f)
            w.writerow(["スライド", "ID", "状態", "元", "新"])
            for row in self.diff_data:
                w.writerow([row["slide"], row.get("id", ""), row["status"], row.get("before", ""), row.get("after", "")])
        messagebox.showinfo("完了", f"CSVを保存しました")


# ============== メインアプリケーション ==============
class InsightSlidesApp:
    def __init__(self, root):
        self.root = root
        self.license_manager = LicenseManager()
        self.config_manager = ConfigManager()
        self.current_mode = "extract"
        self.processing = False
        self.cancel_requested = False
        self.presentation = None
        self.log_buffer = []
        self.extracted_data = []  # グリッド用

        self._setup_window()
        self._apply_styles()
        self._create_menu()
        self._create_layout()
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _setup_window(self):
        tier = self.license_manager.get_tier_info()
        self.root.title(f"{APP_NAME} v{APP_VERSION} - {tier['name']}")
        self.root.geometry("1300x900")
        self.root.minsize(1100, 700)
        self.root.configure(bg=COLOR_PALETTE["bg_primary"])

    def _apply_styles(self):
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure('Main.TFrame', background=COLOR_PALETTE["bg_primary"])
        self.style.configure('Card.TFrame', background=COLOR_PALETTE["bg_elevated"])

    def _create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=t('menu_help'), menu=help_menu)
        help_menu.add_command(label=t('menu_guide'), command=lambda: webbrowser.open(SUPPORT_LINKS["tutorial"]))
        help_menu.add_command(label=t('menu_faq'), command=lambda: webbrowser.open(SUPPORT_LINKS["faq"]))
        help_menu.add_separator()
        help_menu.add_command(label=t('menu_license'), command=self._show_license_dialog)
        help_menu.add_separator()

        lang_menu = tk.Menu(help_menu, tearoff=0)
        help_menu.add_cascade(label=t('lang_menu'), menu=lang_menu)
        lang_menu.add_command(label="English", command=lambda: self._change_language('en'))
        lang_menu.add_command(label="日本語", command=lambda: self._change_language('ja'))

        help_menu.add_separator()
        help_menu.add_command(label=t('menu_about'), command=self._show_about)

    def _create_layout(self):
        if hasattr(self, 'main_container') and self.main_container:
            self.main_container.destroy()

        self.main_container = ttk.Frame(self.root, style='Main.TFrame')
        self.main_container.pack(fill='both', expand=True, padx=SPACING["xl"], pady=SPACING["xl"])
        self.main_container.grid_columnconfigure(0, weight=1)
        self.main_container.grid_rowconfigure(1, weight=1)

        self._create_header(self.main_container)

        content = ttk.Frame(self.main_container, style='Main.TFrame')
        content.grid(row=1, column=0, sticky='nsew')
        content.grid_columnconfigure(1, weight=1)
        content.grid_rowconfigure(0, weight=1)

        self._create_controls(content)
        self._create_output(content)

    def _create_header(self, parent):
        header = tk.Frame(parent, bg=COLOR_PALETTE["bg_elevated"], padx=SPACING["xl"], pady=SPACING["lg"])
        header.grid(row=0, column=0, sticky='ew', pady=(0, SPACING["lg"]))

        # 左: タイトル
        left = tk.Frame(header, bg=COLOR_PALETTE["bg_elevated"])
        left.pack(side='left')

        tk.Label(left, text="◈ Insight Slides", font=FONTS["display"],
                 fg=COLOR_PALETTE["brand_primary"], bg=COLOR_PALETTE["bg_elevated"]).pack(side='left')

        tier = self.license_manager.get_tier_info()
        if tier['name'] != 'Free':
            tk.Label(left, text=f" {tier['name']}", font=FONTS["heading"],
                     fg=COLOR_PALETTE["brand_primary"], bg=COLOR_PALETTE["bg_elevated"]).pack(side='left')

        tk.Label(left, text=f"  {t('app_subtitle')}", font=FONTS["caption"],
                 fg=COLOR_PALETTE["text_muted"], bg=COLOR_PALETTE["bg_elevated"]).pack(side='left', padx=(10, 0))

        # 右: モード + バージョン
        right = tk.Frame(header, bg=COLOR_PALETTE["bg_elevated"])
        right.pack(side='right')

        self.mode_label = tk.Label(right, text=t('mode_extract'), font=FONTS["body"],
                                   fg=COLOR_PALETTE["brand_primary"], bg=COLOR_PALETTE["bg_elevated"])
        self.mode_label.pack(side='left', padx=(0, 20))

        tk.Label(right, text=f"v{APP_VERSION}", font=FONTS["caption"],
                 fg=COLOR_PALETTE["text_muted"], bg=COLOR_PALETTE["bg_elevated"]).pack(side='left')

    def _create_controls(self, parent):
        frame = ttk.Frame(parent, style='Main.TFrame')
        frame.grid(row=0, column=0, sticky='nsew', padx=(0, SPACING["lg"]))
        frame.grid_rowconfigure(4, weight=1)

        # モード切替（3ボタン: 抽出/更新/比較）
        mode_card = ttk.LabelFrame(frame, text="操作モード", padding=SPACING["md"])
        mode_card.grid(row=0, column=0, sticky='ew', pady=(0, SPACING["md"]))
        mode_card.grid_columnconfigure(0, weight=1)
        mode_card.grid_columnconfigure(1, weight=1)
        mode_card.grid_columnconfigure(2, weight=1)

        self.extract_btn = tk.Button(mode_card, text=f"📤 {t('mode_extract_short')}", font=FONTS["body_bold"],
                                     bg=COLOR_PALETTE["brand_primary"], fg="white", relief="flat",
                                     command=self._switch_extract, cursor="hand2")
        self.extract_btn.grid(row=0, column=0, sticky='ew', padx=(0, 3))

        self.update_btn = tk.Button(mode_card, text=f"📥 {t('mode_update_short')}", font=FONTS["body_bold"],
                                    bg=COLOR_PALETTE["bg_secondary"], fg=COLOR_PALETTE["text_primary"], relief="flat",
                                    command=self._switch_update, cursor="hand2")
        self.update_btn.grid(row=0, column=1, sticky='ew', padx=(0, 3))

        self.compare_btn = tk.Button(mode_card, text="🔀 比較", font=FONTS["body_bold"],
                                     bg=COLOR_PALETTE["bg_secondary"], fg=COLOR_PALETTE["text_primary"], relief="flat",
                                     command=self._show_compare_dialog, cursor="hand2")
        self.compare_btn.grid(row=0, column=2, sticky='ew')

        # 説明ラベル
        self.mode_desc_label = tk.Label(mode_card, text="→ PPTXからテキストを抽出して右に表示",
                                        font=FONTS["caption"], fg=COLOR_PALETTE["text_muted"],
                                        bg=COLOR_PALETTE["bg_elevated"])
        self.mode_desc_label.grid(row=1, column=0, columnspan=3, sticky='w', pady=(5, 0))

        # ファイル操作
        self.file_card = ttk.LabelFrame(frame, text=t('panel_file'), padding=SPACING["md"])
        self.file_card.grid(row=1, column=0, sticky='ew', pady=(0, SPACING["md"]))
        self.file_card.grid_columnconfigure(0, weight=1)

        self._create_extract_panel()
        self._create_update_panel()

        # 詳細オプション（折りたたみ）
        self.advanced_var = tk.BooleanVar(value=self.config_manager.get('advanced_expanded', False))
        self.advanced_frame = ttk.LabelFrame(frame, text=f"▶ {t('advanced_options')}", padding=SPACING["md"])
        self.advanced_frame.grid(row=2, column=0, sticky='ew', pady=(0, SPACING["md"]))
        self.advanced_frame.grid_columnconfigure(0, weight=1)
        self.advanced_frame.bind("<Button-1>", self._toggle_advanced)

        self.advanced_content = ttk.Frame(self.advanced_frame)
        self._create_advanced_options()

        if self.advanced_var.get():
            self.advanced_content.grid(row=0, column=0, sticky='ew')
            self.advanced_frame.configure(text=f"▼ {t('advanced_options')}")

        # ステータス
        status_frame = ttk.Frame(frame, style='Main.TFrame')
        status_frame.grid(row=4, column=0, sticky='sew')

        self.status_label = ttk.Label(status_frame, text=t('status_waiting'), font=FONTS["caption"])
        self.status_label.pack(anchor='w')

        self.progress = ttk.Progressbar(status_frame, mode='indeterminate')
        self.progress.pack(fill='x', pady=SPACING["sm"])

        btn_frame = ttk.Frame(status_frame)
        btn_frame.pack(fill='x')
        self.cancel_btn = ttk.Button(btn_frame, text=t('btn_cancel'), command=self._cancel, state='disabled')
        self.cancel_btn.pack(side='left', padx=(0, 5))
        ttk.Button(btn_frame, text=t('btn_clear'), command=self._clear_output).pack(side='left', padx=(0, 5))
        ttk.Button(btn_frame, text=t('btn_copy'), command=self._copy_output).pack(side='left')

        self._switch_extract()

    def _create_extract_panel(self):
        self.extract_frame = ttk.Frame(self.file_card)
        self.extract_frame.grid_columnconfigure(0, weight=1)

        # 出力形式
        fmt_frame = ttk.Frame(self.extract_frame)
        fmt_frame.grid(row=0, column=0, sticky='ew', pady=(0, SPACING["sm"]))
        ttk.Label(fmt_frame, text=t('setting_output_format')).pack(side='left')
        self.output_format_var = tk.StringVar(value=self.config_manager.get('output_format', 'excel'))
        ttk.Combobox(fmt_frame, textvariable=self.output_format_var, values=['excel', 'tab', 'json'],
                     state="readonly", width=12).pack(side='left', padx=5)

        # メタデータ
        self.include_metadata_var = tk.BooleanVar(value=self.config_manager.get('include_metadata', True))
        ttk.Checkbutton(self.extract_frame, text=t('setting_include_meta'),
                        variable=self.include_metadata_var).grid(row=1, column=0, sticky='w')

        # ボタン
        tk.Button(self.extract_frame, text=f"📄 {t('btn_single_file')}", font=FONTS["body"],
                  bg=COLOR_PALETTE["brand_primary"], fg="white", relief="flat",
                  command=self._extract_single).grid(row=2, column=0, sticky='ew', pady=(SPACING["md"], SPACING["sm"]))

        if self.license_manager.can_batch():
            tk.Button(self.extract_frame, text=f"📁 {t('btn_batch_folder')}", font=FONTS["body"],
                      bg=COLOR_PALETTE["bg_secondary"], fg=COLOR_PALETTE["text_primary"], relief="flat",
                      command=self._extract_batch).grid(row=3, column=0, sticky='ew')
        else:
            ttk.Label(self.extract_frame, text=f"📁 {t('btn_batch_folder')} (Standard+)",
                      foreground=COLOR_PALETTE["text_muted"]).grid(row=3, column=0, sticky='w')

    def _create_update_panel(self):
        self.update_frame = ttk.Frame(self.file_card)
        self.update_frame.grid_columnconfigure(0, weight=1)

        limit = self.license_manager.get_update_limit()
        if limit:
            ttk.Label(self.update_frame, text=t('msg_update_limit', limit),
                      foreground=COLOR_PALETTE["warning"]).grid(row=0, column=0, sticky='w', pady=(0, SPACING["sm"]))

        tk.Button(self.update_frame, text=f"📊 {t('btn_from_excel')}", font=FONTS["body"],
                  bg=COLOR_PALETTE["brand_update"], fg="white", relief="flat",
                  command=self._update_excel).grid(row=1, column=0, sticky='ew', pady=(0, SPACING["sm"]))

        tk.Button(self.update_frame, text=f"📄 {t('btn_from_json')}", font=FONTS["body"],
                  bg=COLOR_PALETTE["bg_secondary"], fg=COLOR_PALETTE["text_primary"], relief="flat",
                  command=self._update_json).grid(row=2, column=0, sticky='ew', pady=(0, SPACING["sm"]))

        # Pro機能: 差分プレビュー
        if self.license_manager.is_pro():
            tk.Button(self.update_frame, text=f"👁 {t('btn_diff_preview')}", font=FONTS["body"],
                      bg=COLOR_PALETTE["bg_secondary"], fg=COLOR_PALETTE["text_primary"], relief="flat",
                      command=self._run_preview).grid(row=3, column=0, sticky='ew', pady=(SPACING["sm"], 0))

    def _create_advanced_options(self):
        # スピーカーノート
        self.include_notes_var = tk.BooleanVar(value=False)
        can_notes = self.license_manager.is_pro()
        cb = ttk.Checkbutton(self.advanced_content, text=t('chk_include_notes'),
                             variable=self.include_notes_var,
                             state='normal' if can_notes else 'disabled')
        cb.grid(row=0, column=0, sticky='w')
        if not can_notes:
            ttk.Label(self.advanced_content, text="(Pro)", foreground=COLOR_PALETTE["text_muted"]).grid(row=0, column=1, sticky='w')

        # 自動バックアップ
        self.auto_backup_var = tk.BooleanVar(value=self.config_manager.get('auto_backup', True))
        can_backup = self.license_manager.is_pro()
        cb2 = ttk.Checkbutton(self.advanced_content, text=t('setting_auto_backup'),
                              variable=self.auto_backup_var,
                              state='normal' if can_backup else 'disabled')
        cb2.grid(row=1, column=0, sticky='w')
        if not can_backup:
            ttk.Label(self.advanced_content, text="(Pro)", foreground=COLOR_PALETTE["text_muted"]).grid(row=1, column=1, sticky='w')

    def _toggle_advanced(self, event=None):
        if self.advanced_var.get():
            self.advanced_content.grid_remove()
            self.advanced_frame.configure(text=f"▶ {t('advanced_options')}")
            self.advanced_var.set(False)
        else:
            self.advanced_content.grid(row=0, column=0, sticky='ew')
            self.advanced_frame.configure(text=f"▼ {t('advanced_options')}")
            self.advanced_var.set(True)
        self.config_manager.set('advanced_expanded', self.advanced_var.get())

    def _create_output(self, parent):
        card = ttk.LabelFrame(parent, text=t('panel_output'), padding=SPACING["md"])
        card.grid(row=0, column=1, sticky='nsew')
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(1, weight=1)

        # タブ切替
        self.output_notebook = ttk.Notebook(card)
        self.output_notebook.grid(row=0, column=0, sticky='nsew', rowspan=2)

        # ログタブ
        log_frame = ttk.Frame(self.output_notebook)
        self.output_notebook.add(log_frame, text="📋 ログ")
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(0, weight=1)

        self.output_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, state=tk.DISABLED,
                                                      font=FONTS["body"], bg=COLOR_PALETTE["bg_secondary"])
        self.output_text.grid(row=0, column=0, sticky='nsew')

        # グリッドタブ
        grid_frame = ttk.Frame(self.output_notebook)
        self.output_notebook.add(grid_frame, text="📊 グリッド編集")
        grid_frame.grid_columnconfigure(0, weight=1)
        grid_frame.grid_rowconfigure(0, weight=1)

        self.grid_view = EditableGrid(grid_frame, on_change=self._on_grid_change)
        self.grid_view.grid(row=0, column=0, sticky='nsew')

        # グリッド用ボタン
        grid_btn_frame = ttk.Frame(grid_frame)
        grid_btn_frame.grid(row=1, column=0, sticky='ew', pady=(5, 0))

        tk.Button(grid_btn_frame, text="📥 グリッドから更新適用", font=FONTS["body"],
                  bg=COLOR_PALETTE["brand_update"], fg="white", relief="flat",
                  command=self._apply_grid_to_pptx).pack(side='right')

        ttk.Button(grid_btn_frame, text="📄 Excelエクスポート", command=self._export_grid_excel).pack(side='right', padx=5)

        self._show_welcome()

    def _show_welcome(self):
        tier = self.license_manager.get_tier_info()
        self._update_output(f"{t('welcome_title')}\n{APP_NAME} v{APP_VERSION} ({tier['name']})\n\n", clear=True)

    # === Output helpers ===
    def _update_output(self, text, clear=False):
        self.output_text.configure(state=tk.NORMAL)
        if clear:
            self.output_text.delete('1.0', tk.END)
            self.log_buffer = []
        self.output_text.insert(tk.END, text)
        self.output_text.see(tk.END)
        self.output_text.configure(state=tk.DISABLED)
        self.log_buffer.append(text)

    def _update_output_safe(self, text, clear=False):
        self.root.after(0, lambda: self._update_output(text, clear))

    def _update_status(self, text, color=None):
        self.status_label.configure(text=text)

    def _update_status_safe(self, text, color=None):
        self.root.after(0, lambda: self._update_status(text, color))

    def _log(self, msg, level="info"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        prefix = {"error": "❌ ", "warning": "⚠️ ", "success": "✅ "}.get(level, "")
        self._update_output_safe(f"[{timestamp}] {prefix}{msg}\n")

    def _start_progress(self):
        self.progress.start(10)
        self.processing = True
        self.cancel_requested = False
        self.root.after(0, lambda: self.cancel_btn.configure(state='normal'))

    def _stop_progress(self):
        self.progress.stop()
        self.processing = False
        self.root.after(0, lambda: self.cancel_btn.configure(state='disabled'))

    def _cancel(self):
        if self.processing:
            self.cancel_requested = True
            self._log("キャンセルをリクエスト...", "warning")

    def _clear_output(self):
        self.output_text.configure(state=tk.NORMAL)
        self.output_text.delete('1.0', tk.END)
        self.output_text.configure(state=tk.DISABLED)
        self.log_buffer = []

    def _copy_output(self):
        content = self.output_text.get('1.0', tk.END).strip()
        if content:
            self.root.clipboard_clear()
            self.root.clipboard_append(content)
            messagebox.showinfo(t('dialog_complete'), t('msg_copied'))

    # === Mode switching ===
    def _switch_extract(self):
        self.current_mode = "extract"
        self.mode_label.configure(text=t('mode_extract'), fg=COLOR_PALETTE["brand_primary"])
        self.extract_btn.configure(bg=COLOR_PALETTE["brand_primary"], fg="white")
        self.update_btn.configure(bg=COLOR_PALETTE["bg_secondary"], fg=COLOR_PALETTE["text_primary"])
        self.compare_btn.configure(bg=COLOR_PALETTE["bg_secondary"], fg=COLOR_PALETTE["text_primary"])
        self.mode_desc_label.configure(text="→ PPTXからテキストを抽出して右に表示")
        self.update_frame.grid_remove()
        self.extract_frame.grid(row=0, column=0, sticky='nsew')

    def _switch_update(self):
        self.current_mode = "update"
        self.mode_label.configure(text=t('mode_update'), fg=COLOR_PALETTE["brand_update"])
        self.extract_btn.configure(bg=COLOR_PALETTE["bg_secondary"], fg=COLOR_PALETTE["text_primary"])
        self.update_btn.configure(bg=COLOR_PALETTE["brand_update"], fg="white")
        self.compare_btn.configure(bg=COLOR_PALETTE["bg_secondary"], fg=COLOR_PALETTE["text_primary"])
        self.mode_desc_label.configure(text="→ 右のグリッド内容でPPTXを上書き更新")
        self.extract_frame.grid_remove()
        self.update_frame.grid(row=0, column=0, sticky='nsew')

    def _change_language(self, lang):
        if lang != get_language():
            self.config_manager.set('language', lang)
            set_language(lang)
            self._create_layout()
            messagebox.showinfo(t('dialog_complete'), "言語を変更しました。")

    # === Utility ===
    def clean_text(self, text):
        if text is None:
            return ""
        text = re.sub(r'[\x00-\x08\x0B-\x0C\x0E-\x1F]', '', text)
        text = text.replace('\r\n', '\n').replace('\r', '\n').replace('\v', '\n')
        return text

    def _normalize_for_compare(self, text):
        if text is None:
            return ""
        text = text.replace('\r\n', '\n').replace('\r', '\n').replace('\v', '\n')
        text = text.replace('\u00A0', ' ').replace('\u3000', ' ')
        return text.strip()

    def _texts_are_equal(self, old_text, new_text):
        return self._normalize_for_compare(old_text) == self._normalize_for_compare(new_text)

    def get_shape_type(self, shape):
        try:
            if shape.is_placeholder:
                types_ja = {1: "タイトル", 2: "本文", 3: "図表", 4: "日付", 5: "スライド番号"}
                return types_ja.get(shape.placeholder_format.type, "その他")
            elif hasattr(shape, "has_table") and shape.has_table:
                return "表"
            elif shape.shape_type == 1:
                return "テキストボックス"
            return "その他"
        except:
            return "不明"

    def _create_backup(self, path: str):
        if not self.license_manager.is_pro() or not self.auto_backup_var.get():
            return
        try:
            backup_dir = Path(path).parent / "backup"
            backup_dir.mkdir(exist_ok=True)
            backup_name = f"{Path(path).stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{Path(path).suffix}"
            shutil.copy2(path, backup_dir / backup_name)
            self._log(f"バックアップ作成: {backup_name}")
        except Exception as e:
            self._log(f"バックアップ失敗: {e}", "warning")

    # === Extract ===
    def extract_from_ppt(self, path: str, include_notes: bool = False) -> Tuple[List, Dict]:
        try:
            prs = pptx.Presentation(path)
            data = []
            meta = {'file_name': os.path.basename(path), 'slide_count': len(prs.slides)}

            for slide_num, slide in enumerate(prs.slides, 1):
                if self.cancel_requested:
                    break
                for shape in slide.shapes:
                    try:
                        sid = str(shape.shape_id)
                        stype = self.get_shape_type(shape)

                        if hasattr(shape, "text") and shape.text.strip():
                            data.append({
                                "slide": slide_num, "id": sid, "type": stype, "text": self.clean_text(shape.text)
                            })

                        if hasattr(shape, "has_table") and shape.has_table:
                            for r, row in enumerate(shape.table.rows):
                                for c, cell in enumerate(row.cells):
                                    if cell.text.strip():
                                        data.append({
                                            "slide": slide_num, "id": f"{sid}_t{r}_{c}",
                                            "type": f"表({r+1},{c+1})", "text": self.clean_text(cell.text)
                                        })
                    except:
                        pass

                if include_notes:
                    try:
                        if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
                            notes_text = slide.notes_slide.notes_text_frame.text.strip()
                            if notes_text:
                                data.append({
                                    "slide": slide_num, "id": "notes", "type": t('type_notes'),
                                    "text": self.clean_text(notes_text)
                                })
                    except:
                        pass

            return data, meta
        except Exception as e:
            save_error_log(e, f"extract_from_ppt: {path}")
            self._log(f"読み込みエラー: {e}", "error")
            return [], {}

    def save_to_file(self, data: List[Dict], path: str, fmt: str = "excel") -> bool:
        try:
            if fmt == "excel":
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.append([t('header_slide'), t('header_id'), t('header_type'), t('header_text')])
                for row in data:
                    ws.append([row["slide"], row["id"], row["type"], row["text"]])
                wb.save(path)
            elif fmt == "json":
                with open(path, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
            else:
                with open(path, 'w', encoding='utf-8', newline='') as f:
                    w = csv.writer(f, delimiter='\t')
                    w.writerow([t('header_slide'), t('header_id'), t('header_type'), t('header_text')])
                    for row in data:
                        w.writerow([row["slide"], row["id"], row["type"], row["text"]])
            return True
        except Exception as e:
            save_error_log(e, f"save_to_file: {path}")
            self._log(f"保存エラー: {e}", "error")
            return False

    def _extract_single(self):
        if self.processing:
            return
        path = filedialog.askopenfilename(title="PowerPointファイルを選択", filetypes=[("PowerPoint", "*.pptx")])
        if not path:
            return
        self.config_manager.set('last_directory', os.path.dirname(path))

        include_notes = self.include_notes_var.get() if self.license_manager.is_pro() else False

        def run():
            try:
                self._start_progress()
                self._update_status_safe("処理中...")
                self._update_output_safe(f"\n📄 処理開始: {os.path.basename(path)}\n", clear=True)

                data, meta = self.extract_from_ppt(path, include_notes)
                if self.cancel_requested:
                    return self._log("キャンセルされました", "warning")

                if data:
                    # グリッドにロード
                    self.extracted_data = data
                    self.root.after(0, lambda: self.grid_view.load_data(data))
                    self.root.after(0, lambda: self.output_notebook.select(1))  # グリッドタブへ

                    # ファイル保存
                    fmt = self.output_format_var.get()
                    ext = {"excel": ".xlsx", "json": ".json"}.get(fmt, ".txt")
                    out = os.path.splitext(path)[0] + "_抽出" + ext
                    if self.save_to_file(data, out, fmt):
                        self._log(f"✅ 抽出完了: {len(data)}件 → {os.path.basename(out)}", "success")
                        self._update_status_safe(f"完了: {len(data)}件")
                else:
                    self._log("テキストが見つかりませんでした", "warning")
            except Exception as e:
                save_error_log(e, "_extract_single")
                self._log(f"エラー: {e}", "error")
            finally:
                self._stop_progress()

        threading.Thread(target=run, daemon=True).start()

    def _extract_batch(self):
        if self.processing:
            return
        folder = filedialog.askdirectory(title="フォルダを選択")
        if not folder:
            return

        include_notes = self.include_notes_var.get() if self.license_manager.is_pro() else False

        def run():
            try:
                self._start_progress()
                self._update_output_safe(f"\n📁 フォルダ処理: {folder}\n", clear=True)

                files = [f for f in Path(folder).glob("*.pptx") if not f.name.startswith("~$")]
                if not files:
                    return self._log("PPTXファイルが見つかりません", "warning")

                self._log(f"発見: {len(files)}件")
                total = 0

                for i, f in enumerate(files, 1):
                    if self.cancel_requested:
                        break
                    self._log(f"[{i}/{len(files)}] {f.name}")
                    data, meta = self.extract_from_ppt(str(f), include_notes)
                    if data:
                        fmt = self.output_format_var.get()
                        ext = {"excel": ".xlsx", "json": ".json"}.get(fmt, ".txt")
                        out = str(f.with_suffix('')) + "_抽出" + ext
                        self.save_to_file(data, out, fmt)
                        total += len(data)

                self._log(f"✅ バッチ抽出完了: {total}件", "success")
            except Exception as e:
                self._log(f"エラー: {e}", "error")
            finally:
                self._stop_progress()

        threading.Thread(target=run, daemon=True).start()

    # === Update ===
    def _load_updates(self, path: str, source: str) -> Dict:
        updates = {}
        try:
            if source == "excel":
                wb = openpyxl.load_workbook(path)
                ws = wb.active
                headers = [c.value for c in ws[1]]
                try:
                    si = headers.index("スライド番号") if "スライド番号" in headers else headers.index("slide")
                    oi = headers.index("オブジェクトID") if "オブジェクトID" in headers else headers.index("id")
                    ti = headers.index("テキスト内容") if "テキスト内容" in headers else headers.index("text")
                except:
                    self._log("ヘッダー形式が不正です", "error")
                    return {}
                for row in list(ws.rows)[1:]:
                    try:
                        sn = int(row[si].value) if row[si].value else None
                        oid = str(row[oi].value) if row[oi].value else None
                        txt = str(row[ti].value) if row[ti].value else ""
                        if txt == "None":
                            txt = ""
                        if sn and oid:
                            updates[(sn, oid)] = txt
                    except:
                        pass
            elif source == "json":
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                for item in data:
                    sn = item.get('スライド番号') or item.get('slide')
                    oid = item.get('オブジェクトID') or item.get('id')
                    txt = item.get('テキスト内容') or item.get('text', '')
                    if sn and oid:
                        updates[(int(sn), str(oid))] = str(txt)
        except Exception as e:
            self._log(f"読み込みエラー: {e}", "error")
        return updates

    def _update_ppt(self, ppt_path: str, updates: Dict, preview: bool = False) -> Tuple[int, int, List]:
        limit = self.license_manager.get_update_limit()
        self.presentation = pptx.Presentation(ppt_path)
        updated, skipped = 0, 0
        changes = []

        for slide_idx, slide in enumerate(self.presentation.slides, 1):
            if self.cancel_requested:
                break
            if limit and slide_idx > limit:
                skipped += len([k for k in updates if k[0] == slide_idx])
                continue

            for shape in slide.shapes:
                try:
                    sid = str(shape.shape_id)
                    key = (slide_idx, sid)

                    if key in updates and hasattr(shape, "text"):
                        new_txt = updates[key]
                        old_txt = shape.text
                        if not self._texts_are_equal(old_txt, new_txt):
                            changes.append({'slide': slide_idx, 'id': sid, 'old': old_txt[:50], 'new': new_txt[:50]})
                            if not preview:
                                shape.text = self._normalize_for_compare(new_txt)
                                updated += 1

                    if hasattr(shape, "has_table") and shape.has_table:
                        for r, row in enumerate(shape.table.rows):
                            for c, cell in enumerate(row.cells):
                                ckey = (slide_idx, f"{sid}_t{r}_{c}")
                                if ckey in updates:
                                    new_txt = updates[ckey]
                                    old_txt = cell.text
                                    if not self._texts_are_equal(old_txt, new_txt):
                                        changes.append({'slide': slide_idx, 'id': ckey[1], 'old': old_txt[:30], 'new': new_txt[:30]})
                                        if not preview:
                                            cell.text = self._normalize_for_compare(new_txt)
                                            updated += 1
                except Exception as e:
                    skipped += 1

        return updated, skipped, changes

    def _run_update(self, source: str):
        if self.processing:
            return

        limit = self.license_manager.get_update_limit()
        if limit:
            if not messagebox.askyesno("確認", f"Free版では最初の{limit}スライドのみ更新されます。続行しますか？"):
                return

        ftypes = [("Excel", "*.xlsx")] if source == "excel" else [("JSON", "*.json")]
        data_path = filedialog.askopenfilename(title="編集済みファイルを選択", filetypes=ftypes)
        if not data_path:
            return
        ppt_path = filedialog.askopenfilename(title="PowerPointファイルを選択", filetypes=[("PowerPoint", "*.pptx")])
        if not ppt_path:
            return

        def run():
            try:
                self._start_progress()
                self._update_output_safe(f"\n📥 更新処理開始\n", clear=True)
                self._create_backup(ppt_path)

                updates = self._load_updates(data_path, source)
                if not updates:
                    return self._log("更新データなし", "warning")

                self._log(f"読み込み: {len(updates)}件")
                updated, skipped, _ = self._update_ppt(ppt_path, updates)

                if self.cancel_requested:
                    return

                def save():
                    out = filedialog.asksaveasfilename(defaultextension=".pptx", filetypes=[("PowerPoint", "*.pptx")],
                                                       initialfile=os.path.splitext(os.path.basename(ppt_path))[0] + "_更新済み.pptx")
                    if out:
                        self.presentation.save(out)
                        self._log(f"✅ 保存完了: {os.path.basename(out)}", "success")
                        messagebox.showinfo("完了", f"更新: {updated}件\nスキップ: {skipped}件")

                self.root.after(0, save)
            except Exception as e:
                self._log(f"エラー: {e}", "error")
            finally:
                self._stop_progress()

        threading.Thread(target=run, daemon=True).start()

    def _update_excel(self):
        self._run_update("excel")

    def _update_json(self):
        self._run_update("json")

    def _run_preview(self):
        data_path = filedialog.askopenfilename(title="編集済みファイルを選択", filetypes=[("Excel/TXT", "*.xlsx *.txt")])
        if not data_path:
            return
        ppt_path = filedialog.askopenfilename(title="PowerPointファイルを選択", filetypes=[("PowerPoint", "*.pptx")])
        if not ppt_path:
            return

        def run():
            try:
                self._start_progress()
                self._update_output_safe(f"\n👁 差分プレビュー\n", clear=True)
                source = "excel" if data_path.endswith('.xlsx') else "json"
                updates = self._load_updates(data_path, source)
                if not updates:
                    return self._log("更新データなし", "warning")

                _, _, changes = self._update_ppt(ppt_path, updates, preview=True)
                if changes:
                    self._log(f"\n変更箇所: {len(changes)}件")
                    for i, c in enumerate(changes[:20], 1):
                        self._update_output_safe(f"[{i}] スライド{c['slide']} ID:{c['id']}\n  旧: {c['old']}\n  新: {c['new']}\n\n")
                else:
                    self._log("変更箇所なし")
            except Exception as e:
                self._log(f"エラー: {e}", "error")
            finally:
                self._stop_progress()

        threading.Thread(target=run, daemon=True).start()

    # === 比較機能 ===
    def _show_compare_dialog(self):
        CompareDialog(self.root, self._run_compare)

    def _run_compare(self, file1: str, file2: str, ignore_ws: bool):
        def run():
            try:
                self._start_progress()
                self._update_output_safe(f"\n🔀 比較処理中...\n", clear=True)

                data1, _ = self.extract_from_ppt(file1)
                data2, _ = self.extract_from_ppt(file2)

                # マッピング
                map1 = {(d["slide"], d["id"]): d["text"] for d in data1}
                map2 = {(d["slide"], d["id"]): d["text"] for d in data2}

                all_keys = set(map1.keys()) | set(map2.keys())
                diff_data = []
                stats = {"same": 0, "changed": 0, "added": 0, "removed": 0}

                for key in sorted(all_keys):
                    t1 = map1.get(key)
                    t2 = map2.get(key)

                    if t1 and t2:
                        if self._texts_are_equal(t1, t2):
                            status = "一致"
                            stats["same"] += 1
                        else:
                            status = "変更"
                            stats["changed"] += 1
                    elif t1:
                        status = "削除"
                        stats["removed"] += 1
                    else:
                        status = "追加"
                        stats["added"] += 1

                    diff_data.append({
                        "slide": key[0], "id": key[1], "status": status,
                        "before": t1 or "", "after": t2 or ""
                    })

                self._log(f"比較完了: 一致{stats['same']} 変更{stats['changed']} 追加{stats['added']} 削除{stats['removed']}")

                self.root.after(0, lambda: CompareResultWindow(
                    self.root, os.path.basename(file1), os.path.basename(file2),
                    diff_data, stats, on_apply=self._apply_compare_result
                ))
            except Exception as e:
                self._log(f"エラー: {e}", "error")
            finally:
                self._stop_progress()

        threading.Thread(target=run, daemon=True).start()

    def _apply_compare_result(self, selected_data: List[Dict]):
        # 比較結果をグリッドに反映
        grid_data = []
        for item in selected_data:
            grid_data.append({
                "slide": item["slide"], "id": item.get("id", ""),
                "type": "", "text": item["text"]
            })
        self.grid_view.load_data(grid_data)
        self.output_notebook.select(1)

    # === グリッド操作 ===
    def _on_grid_change(self, item, column, value):
        pass  # 変更時の追加処理があれば

    def _apply_grid_to_pptx(self):
        if not self.grid_view.get_data():
            messagebox.showwarning("警告", "グリッドにデータがありません")
            return

        ppt_path = filedialog.askopenfilename(title="更新するPowerPointを選択", filetypes=[("PowerPoint", "*.pptx")])
        if not ppt_path:
            return

        # グリッドデータから更新辞書作成
        grid_data = self.grid_view.get_data()
        updates = {}
        for row in grid_data:
            try:
                sn = int(row["slide"])
                oid = row["id"]
                txt = row["text"]
                if sn and oid:
                    updates[(sn, oid)] = txt
            except:
                pass

        if not updates:
            messagebox.showwarning("警告", "有効な更新データがありません")
            return

        def run():
            try:
                self._start_progress()
                self._log(f"グリッドから更新: {len(updates)}件")
                self._create_backup(ppt_path)

                updated, skipped, _ = self._update_ppt(ppt_path, updates)

                def save():
                    out = filedialog.asksaveasfilename(defaultextension=".pptx", filetypes=[("PowerPoint", "*.pptx")],
                                                       initialfile=os.path.splitext(os.path.basename(ppt_path))[0] + "_更新済み.pptx")
                    if out:
                        self.presentation.save(out)
                        self._log(f"✅ 保存完了: {out}", "success")
                        messagebox.showinfo("完了", f"更新: {updated}件")

                self.root.after(0, save)
            except Exception as e:
                self._log(f"エラー: {e}", "error")
            finally:
                self._stop_progress()

        threading.Thread(target=run, daemon=True).start()

    def _export_grid_excel(self):
        data = self.grid_view.get_data()
        if not data:
            messagebox.showwarning("警告", "エクスポートするデータがありません")
            return

        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not path:
            return

        if self.save_to_file(data, path, "excel"):
            messagebox.showinfo("完了", f"エクスポート完了: {path}")

    # === Dialogs ===
    def _show_license_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title(t('license_title'))
        dialog.geometry("450x400")
        dialog.transient(self.root)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding=20)
        frame.pack(fill='both', expand=True)

        tier = self.license_manager.get_tier_info()
        ttk.Label(frame, text=t('license_current'), font=FONTS["heading"]).pack(anchor='w')
        ttk.Label(frame, text=f"{tier['badge']} ({tier['name']})", font=FONTS["body_bold"]).pack(anchor='w', pady=(5, 15))

        ttk.Label(frame, text="形式: INS-SLIDE-{TIER}-XXXX-XXXX-CC", font=FONTS["small"],
                  foreground=COLOR_PALETTE["text_muted"]).pack(anchor='w')

        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=15)

        ttk.Label(frame, text=t('license_enter_key')).pack(anchor='w')
        key_var = tk.StringVar()
        ttk.Entry(frame, textvariable=key_var, width=40, font=FONTS["body"]).pack(fill='x', pady=(5, 15))

        def activate():
            ok, msg = self.license_manager.activate(key_var.get())
            if ok:
                messagebox.showinfo(t('dialog_complete'), msg)
                dialog.destroy()
                self._create_layout()
            else:
                messagebox.showerror(t('dialog_error'), msg)

        def deactivate():
            self.license_manager.deactivate()
            messagebox.showinfo(t('dialog_complete'), t('license_deactivated'))
            dialog.destroy()
            self._create_layout()

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill='x')

        if tier['name'] != 'Free':
            ttk.Button(btn_frame, text=t('btn_deactivate'), command=deactivate).pack(side='left')

        ttk.Button(btn_frame, text=t('btn_activate'), command=activate).pack(side='left', padx=5)
        ttk.Button(btn_frame, text=t('btn_close'), command=dialog.destroy).pack(side='right')

    def _show_about(self):
        tier = self.license_manager.get_tier_info()
        messagebox.showinfo(t('menu_about'),
            f"{APP_NAME} v{APP_VERSION}\n\n"
            f"ライセンス: {tier['name']}\n\n"
            f"統一ライセンス形式:\n"
            f"INS-SLIDE-{{TIER}}-XXXX-XXXX-CC\n\n"
            f"by Harmonic Insight\n© 2025"
        )

    def _on_closing(self):
        if self.processing:
            if not messagebox.askokcancel("確認", "処理中です。終了しますか？"):
                return
        self.root.destroy()


def main():
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass
    root = tk.Tk()
    InsightSlidesApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
