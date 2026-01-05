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
import sys
import re
import csv
import json
import hashlib
import random
import string
import webbrowser
import traceback
from datetime import datetime, timedelta

# insight-common ライセンスモジュールをインポート
import importlib.util
_license_module_path = os.path.join(os.path.dirname(__file__), 'insight-common', 'license', 'python', '__init__.py')
_spec = importlib.util.spec_from_file_location("insight_license", _license_module_path)
_license_module = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_license_module)
ProductCode = _license_module.ProductCode
InsightLicenseTier = _license_module.LicenseTier
LicenseInfo = _license_module.LicenseInfo
LicenseValidator = _license_module.LicenseValidator
get_feature_limits = _license_module.get_feature_limits
TIER_NAMES = _license_module.TIER_NAMES
INSIGHT_TIERS = _license_module.TIERS
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
        'panel_input': 'Input (Single File)',
        'panel_output_file': 'Output (Single File)',
        'panel_settings': 'Settings',
        'panel_status': 'Status',
        'panel_output': 'Extracted Data',
        'panel_extract_options': 'Extract Options',
        'panel_update_options': 'Update Options',
        'panel_extract_run': 'Run Extract',
        'panel_update_run': 'Run Update',
        'panel_pro_features': 'Pro Features',
        'btn_load_pptx': 'Load PPTX',
        'btn_load_excel': 'Load Excel',
        'btn_load_json': 'Load JSON',
        'btn_single_file': 'Select File',
        'btn_from_excel': 'From Excel',
        'btn_from_json': 'From JSON',
        'btn_apply_pptx': 'Apply to PPTX',
        'btn_export_to_excel': 'Export to Excel',
        'btn_export_to_json': 'Export to JSON',
        'panel_batch': 'Folder Batch',
        'btn_batch_extract': 'Batch Extract',
        'btn_batch_update': 'Batch Update',
        'btn_batch_export_excel': 'Export to Folder (Excel)',
        'btn_batch_export_json': 'Export to Folder (JSON)',
        'btn_batch_import_excel': 'Import from Folder (Excel)',
        'btn_batch_import_json': 'Import from Folder (JSON)',
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
        # UI elements
        'mode_section': 'Mode',
        'btn_compare': '2-File Compare',
        'show_detail': 'Show details',
        'welcome_guide_title': 'Edit PowerPoint Text',
        'guide_step1': 'Select a PPTX file from the left panel',
        'guide_step2': 'Text will be displayed in a list',
        'guide_step3': 'Double-click a cell to edit',
        'guide_step4': 'Click "Apply" to update PPTX',
        'btn_apply': 'Apply',
        'btn_export_excel': 'Excel Export',
        'btn_export_json': 'JSON Export',
        'filter_label': 'Filter:',
        'mode_desc_extract': 'Extract text from PPTX for editing',
        'mode_desc_update': 'Apply edited data to PPTX',
        # Grid toolbar
        'btn_clear_grid': 'Clear',
        'btn_replace_all': 'Replace All',
        'btn_undo': 'Undo',
        'btn_redo': 'Redo',
        # Replace dialog
        'replace_search': 'Search:',
        'replace_with': 'Replace:',
        'btn_replace': 'Replace',
        # Compare dialog
        'compare_title': 'Compare 2 PowerPoint files',
        'compare_file1': 'Original:',
        'compare_file2': 'New file:',
        'btn_browse': 'Browse',
        'compare_ignore_ws': 'Ignore whitespace',
        'btn_run_compare': 'Compare',
        # Compare result
        'btn_export_csv': 'CSV Export',
        'header_select': 'Select',
        'header_status': 'Status',
        'btn_select_original': 'All Original',
        'btn_select_new': 'All New',
        'btn_apply_selection': 'Apply Selection',
        # Log dialog
        'btn_copy_log': 'Copy',
        'btn_clear_log': 'Clear',
        # License dialog (auth)
        'license_auth_title': 'License Activation',
        'license_email': 'Email Address:',
        'license_key': 'License Key:',
        'license_wrong_product': 'This license key is not valid for Insight Slides',
        'license_perpetual': 'Perpetual',
        'license_expiry_warning': 'Your license will expire in {0} days ({1}). Please renew.',
        'license_expired': 'Your license has expired. Please renew to continue using all features.',
        'license_trial_link': 'Request Trial',
        'license_email_required': 'Please enter your email address',
        'license_status_active': 'Active',
        'license_status_expired': 'Expired',
        'license_valid_until': 'Valid until: {0}',
        'license_days_remaining': '({0} days remaining)',
        'license_feature_restricted': 'This feature requires a Pro license. Current: {0}',
        'license_batch_restricted': 'Batch processing requires a Pro license.',
        'license_json_restricted': 'JSON export requires a Pro license.',
        'license_continue_free': 'Continue as Free',
        # Status messages
        'status_slides_items': '{0} slides / {1} items',
        'status_complete_items': 'Complete: {0} items',
        'status_batch_complete': 'Batch extract complete: {0} items ({1})',
        'status_update_complete': 'Update complete: {0} items',
        'lang_changed': 'Language changed.',
        # Log messages
        'log_cancelled': 'Cancelled',
        'log_cancel_request': 'Cancellation requested...',
        'log_no_text': 'No text found',
        'log_error': 'Error: {0}',
        'log_found_files': 'Found: {0} files',
        'log_no_pptx_found': 'No PPTX files found',
        'log_invalid_header': 'Invalid header format',
        'log_no_update_data': 'No update data',
        'log_processing': 'Processing...',
        'dialog_select_folder': 'Select Folder (containing PPTX files)',
        'dialog_select_folder_update': 'Select Folder (*_extracted{0} + PPTX)',
        'dialog_select_pptx': 'Select PowerPoint to update',
        'dialog_processing_exit': 'Processing in progress. Exit anyway?',
        'dialog_confirm_title': 'Confirm',
        'result_updated': 'Updated: {0} items\nSkipped: {1} items',
        'result_replaced': '{0} items replaced',
        'result_applied': '{0} items applied',
        'result_csv_saved': 'CSV saved',
        'result_export_complete': 'Export complete: {0}',
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
        'panel_input': '入力（1ファイル）',
        'panel_output_file': '出力（1ファイル）',
        'panel_settings': '処理設定',
        'panel_status': '処理状況',
        'panel_output': '抽出結果',
        'panel_extract_options': '抽出オプション',
        'panel_update_options': '更新オプション',
        'panel_extract_run': '抽出実行',
        'panel_update_run': '更新実行',
        'panel_pro_features': '拡張機能',
        'btn_load_pptx': 'PPTX読込',
        'btn_load_excel': 'Excel読込',
        'btn_load_json': 'JSON読込',
        'btn_single_file': 'ファイル選択',
        'btn_from_excel': 'Excelから更新',
        'btn_from_json': 'JSONから更新',
        'btn_apply_pptx': 'PPTXに反映',
        'btn_export_to_excel': 'Excel出力',
        'btn_export_to_json': 'JSON出力',
        'panel_batch': 'フォルダ一括',
        'btn_batch_extract': '一括抽出',
        'btn_batch_update': '一括更新',
        'btn_batch_export_excel': 'フォルダに出力 (Excel)',
        'btn_batch_export_json': 'フォルダに出力 (JSON)',
        'btn_batch_import_excel': 'フォルダから読込 (Excel)',
        'btn_batch_import_json': 'フォルダから読込 (JSON)',
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
        # UI elements
        'mode_section': '操作モード',
        'btn_compare': '2ファイル比較',
        'show_detail': '詳細を表示',
        'welcome_guide_title': 'PowerPointテキストを編集',
        'guide_step1': '左のパネルでPPTXファイルを選択',
        'guide_step2': 'テキストが一覧で表示されます',
        'guide_step3': 'セルをダブルクリックして編集',
        'guide_step4': '「更新を適用」でPPTXに反映',
        'btn_apply': '更新を適用',
        'btn_export_excel': 'Excelエクスポート',
        'btn_export_json': 'JSONエクスポート',
        'filter_label': 'フィルタ:',
        'mode_desc_extract': 'PPTXからテキストを抽出して編集',
        'mode_desc_update': '編集したデータをPPTXに反映',
        # Grid toolbar
        'btn_clear_grid': 'クリア',
        'btn_replace_all': '一括置換',
        'btn_undo': '元に戻す',
        'btn_redo': 'やり直し',
        # Replace dialog
        'replace_search': '検索:',
        'replace_with': '置換:',
        'btn_replace': '置換',
        # Compare dialog
        'compare_title': '2つのPowerPointファイルを比較',
        'compare_file1': '元ファイル:',
        'compare_file2': '新ファイル:',
        'btn_browse': '参照',
        'compare_ignore_ws': '空白の違いを無視',
        'btn_run_compare': '比較実行',
        # Compare result
        'btn_export_csv': 'CSVエクスポート',
        'header_select': '採用',
        'header_status': '状態',
        'btn_select_original': '全て元',
        'btn_select_new': '全て新',
        'btn_apply_selection': '選択を反映',
        # Log dialog
        'btn_copy_log': 'コピー',
        'btn_clear_log': 'クリア',
        # License dialog (auth)
        'license_auth_title': 'ライセンス認証',
        'license_email': 'メールアドレス:',
        'license_key': 'ライセンスキー:',
        'license_wrong_product': 'このライセンスキーはInsight Slidesには適用できません',
        'license_perpetual': '永続',
        'license_expiry_warning': 'ライセンスの有効期限まであと{0}日です（{1}まで）。更新をご検討ください。',
        'license_expired': 'ライセンスの有効期限が切れました。継続してご利用いただくには更新が必要です。',
        'license_trial_link': 'トライアル申請',
        'license_email_required': 'メールアドレスを入力してください',
        'license_status_active': '有効',
        'license_status_expired': '期限切れ',
        'license_valid_until': '有効期限: {0}',
        'license_days_remaining': '（残り{0}日）',
        'license_feature_restricted': 'この機能はProライセンスが必要です。現在: {0}',
        'license_batch_restricted': 'フォルダ一括処理はProライセンスが必要です。',
        'license_json_restricted': 'JSON出力はProライセンスが必要です。',
        'license_continue_free': 'Free版で続行',
        # Status messages
        'status_slides_items': '{0}スライド / {1}項目',
        'status_complete_items': '完了: {0}件',
        'status_batch_complete': 'バッチ抽出完了: {0}件 ({1})',
        'status_update_complete': '更新完了: {0}件',
        'lang_changed': '言語を変更しました。',
        # Log messages
        'log_cancelled': 'キャンセルされました',
        'log_cancel_request': 'キャンセルをリクエスト...',
        'log_no_text': 'テキストが見つかりませんでした',
        'log_error': 'エラー: {0}',
        'log_found_files': '発見: {0}件',
        'log_no_pptx_found': 'PPTXファイルが見つかりません',
        'log_invalid_header': 'ヘッダー形式が不正です',
        'log_no_update_data': '更新データなし',
        'log_processing': '処理中...',
        'dialog_select_folder': 'フォルダを選択 (PPTXファイルを含む)',
        'dialog_select_folder_update': 'フォルダを選択 (*_抽出{0} + PPTX)',
        'dialog_select_pptx': '更新するPowerPointを選択',
        'dialog_processing_exit': '処理中です。終了しますか？',
        'dialog_confirm_title': '確認',
        'result_updated': '更新: {0}件\nスキップ: {1}件',
        'result_replaced': '{0} 件を置換しました',
        'result_applied': '{0} 件を反映しました',
        'result_csv_saved': 'CSVを保存しました',
        'result_export_complete': 'エクスポート完了: {0}',
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


# ============== ライセンス設定（insight-common 統合） ==============
PRODUCT_CODE = ProductCode.SLIDE
EXPIRY_WARNING_DAYS = 30  # 期限切れ警告の日数

# ローカルティア定義（FREE追加）
class LicenseTier:
    FREE = "FREE"
    TRIAL = "TRIAL"
    STD = "STD"
    PRO = "PRO"
    ENT = "ENT"

# ティア別設定（InsightSlide固有）
# json: 1ファイルJSON入出力, batch: フォルダ一括処理, compare: 2ファイル比較
TIERS = {
    LicenseTier.FREE: {'name': 'Free', 'name_ja': '無料版', 'badge': 'Free', 'update_limit': 3, 'batch': False, 'json': False, 'compare': False},
    LicenseTier.TRIAL: {'name': 'Trial', 'name_ja': 'トライアル', 'badge': 'Trial', 'update_limit': None, 'batch': True, 'json': True, 'compare': True},
    LicenseTier.STD: {'name': 'Standard', 'name_ja': 'スタンダード', 'badge': 'Standard', 'update_limit': None, 'batch': False, 'json': False, 'compare': True},
    LicenseTier.PRO: {'name': 'Professional', 'name_ja': 'プロフェッショナル', 'badge': 'Pro', 'update_limit': None, 'batch': True, 'json': True, 'compare': True},
    LicenseTier.ENT: {'name': 'Enterprise', 'name_ja': 'エンタープライズ', 'badge': 'Enterprise', 'update_limit': None, 'batch': True, 'json': True, 'compare': True},
}


class LicenseManager:
    """insight-common 統合ライセンスマネージャー"""

    def __init__(self):
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        self.validator = LicenseValidator()
        self.license_info: Dict = {}
        self.insight_info: Optional[LicenseInfo] = None
        self._load_license()

    def _load_license(self):
        """保存されたライセンス情報を読み込む"""
        self.license_info = {'type': LicenseTier.FREE, 'key': '', 'email': '', 'expires': None}

        if LICENSE_FILE.exists():
            try:
                with open(LICENSE_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)

                if data.get('key'):
                    # 有効期限を復元
                    expires_at = None
                    if data.get('expires'):
                        try:
                            expires_at = datetime.fromisoformat(data['expires'])
                        except:
                            pass

                    # insight-common で検証
                    self.insight_info = self.validator.validate(data['key'], expires_at)

                    if self.insight_info.is_valid:
                        # 製品チェック
                        if self.validator.is_product_covered(self.insight_info, PRODUCT_CODE):
                            tier = self._map_insight_tier(self.insight_info.tier)
                            self.license_info = {
                                'type': tier,
                                'key': data['key'],
                                'email': data.get('email', ''),
                                'expires': data.get('expires')
                            }
                            return

            except Exception as e:
                print(f"License load error: {e}")

    def _save_license(self):
        """ライセンス情報を保存"""
        with open(LICENSE_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.license_info, f, ensure_ascii=False, indent=2)

    def _map_insight_tier(self, tier: Optional[InsightLicenseTier]) -> str:
        """insight-common のティアをローカルティアにマップ"""
        if not tier:
            return LicenseTier.FREE
        mapping = {
            InsightLicenseTier.TRIAL: LicenseTier.TRIAL,
            InsightLicenseTier.STD: LicenseTier.STD,
            InsightLicenseTier.PRO: LicenseTier.PRO,
            InsightLicenseTier.ENT: LicenseTier.ENT,
        }
        return mapping.get(tier, LicenseTier.FREE)

    def activate(self, email: str, key: str) -> Tuple[bool, str]:
        """ライセンスをアクティベート"""
        if not email or not key:
            return False, t('license_enter_prompt')

        # insight-common で検証
        self.insight_info = self.validator.validate(key.strip())

        if not self.insight_info.is_valid:
            error_msg = self.insight_info.error or t('license_invalid')
            return False, error_msg

        # 製品チェック
        if not self.validator.is_product_covered(self.insight_info, PRODUCT_CODE):
            return False, t('license_wrong_product')

        # 有効期限を計算（初回アクティベーション時）
        expires_str = None
        if self.insight_info.tier and self.insight_info.tier != InsightLicenseTier.ENT:
            tier_config = INSIGHT_TIERS.get(self.insight_info.tier, {})
            duration_months = tier_config.get('duration_months')
            duration_days = tier_config.get('duration_days')

            if duration_days:
                expires = datetime.now() + timedelta(days=duration_days)
                expires_str = expires.isoformat()
            elif duration_months:
                now = datetime.now()
                new_month = now.month + duration_months
                new_year = now.year + (new_month - 1) // 12
                new_month = (new_month - 1) % 12 + 1
                expires = datetime(new_year, new_month, min(now.day, 28))
                expires_str = expires.isoformat()

        tier = self._map_insight_tier(self.insight_info.tier)
        self.license_info = {
            'type': tier,
            'key': key.strip().upper(),
            'email': email.strip(),
            'expires': expires_str
        }
        self._save_license()

        tier_info = TIERS.get(tier, TIERS[LicenseTier.FREE])
        name = tier_info['name_ja'] if get_language() == 'ja' else tier_info['name']
        return True, t('license_activated', name)

    def deactivate(self):
        """ライセンスを解除"""
        self.license_info = {'type': LicenseTier.FREE, 'key': '', 'email': '', 'expires': None}
        self.insight_info = None
        if LICENSE_FILE.exists():
            LICENSE_FILE.unlink()

    def get_tier(self) -> str:
        return self.license_info.get('type', LicenseTier.FREE)

    def get_tier_info(self) -> Dict:
        return TIERS.get(self.get_tier(), TIERS[LicenseTier.FREE])

    def get_update_limit(self) -> Optional[int]:
        return self.get_tier_info().get('update_limit')

    def can_batch(self) -> bool:
        """フォルダ一括処理が可能か（PRO/Trial/ENT）"""
        return self.get_tier_info().get('batch', False)

    def can_json(self) -> bool:
        """1ファイルJSON入出力が可能か（PRO/Trial/ENT）"""
        return self.get_tier_info().get('json', False)

    def can_compare(self) -> bool:
        """2ファイル比較が可能か（STD以上）"""
        return self.get_tier_info().get('compare', False)

    def is_pro(self) -> bool:
        """後方互換性のため維持（PRO機能 = batch + json）"""
        return self.can_batch() and self.can_json()

    def is_activated(self) -> bool:
        """ライセンスがアクティベートされているか"""
        return self.get_tier() != LicenseTier.FREE

    def get_days_until_expiry(self) -> Optional[int]:
        """有効期限までの日数を取得（期限なしの場合はNone）"""
        expires_str = self.license_info.get('expires')
        if not expires_str:
            return None
        try:
            expires = datetime.fromisoformat(expires_str)
            delta = expires - datetime.now()
            return delta.days
        except:
            return None

    def should_show_expiry_warning(self) -> bool:
        """期限切れ警告を表示すべきか"""
        days = self.get_days_until_expiry()
        if days is None:
            return False
        return 0 < days <= EXPIRY_WARNING_DAYS

    def get_expiry_date_str(self) -> str:
        """有効期限の表示文字列"""
        expires_str = self.license_info.get('expires')
        if not expires_str:
            return t('license_perpetual') if self.get_tier() == LicenseTier.ENT else '-'
        try:
            expires = datetime.fromisoformat(expires_str)
            return expires.strftime('%Y/%m/%d')
        except:
            return '-'


# ============== モダンデザインシステム ==============
# B2B SaaS品質 - Notion/Linear/Figma風

# カラーパレット（洗練されたニュートラル + 落ち着いたブルー）
COLOR_PALETTE = {
    # 背景
    "bg_primary": "#FFFFFF",       # メイン背景
    "bg_secondary": "#F8FAFC",     # セカンダリ背景（カード内）
    "bg_elevated": "#F1F5F9",      # 強調背景（ホバー等）
    "bg_sidebar": "#FAFBFC",       # サイドバー背景
    "bg_card": "#FFFFFF",          # カード背景
    "bg_input": "#FFFFFF",         # 入力フィールド背景

    # テキスト（4段階の階層）
    "text_primary": "#1F2937",     # メインテキスト（見出し）
    "text_secondary": "#374151",   # 本文テキスト
    "text_tertiary": "#6B7280",    # 補助テキスト
    "text_muted": "#9CA3AF",       # 薄いテキスト（注釈）
    "text_placeholder": "#D1D5DB", # プレースホルダー

    # ブランドカラー（落ち着いたブルー系）
    "brand_primary": "#2563EB",    # プライマリブルー
    "brand_hover": "#1D4ED8",      # ホバー時（濃い）
    "brand_light": "#DBEAFE",      # 薄いブルー（選択背景）
    "brand_muted": "#93C5FD",      # ミュートブルー

    # セカンダリアクション
    "secondary_default": "#F3F4F6",  # セカンダリボタン背景
    "secondary_hover": "#E5E7EB",    # セカンダリホバー
    "secondary_border": "#D1D5DB",   # セカンダリボーダー

    # 機能別カラー
    "action_update": "#059669",    # 更新（グリーン）
    "action_compare": "#7C3AED",   # 比較（パープル）
    "action_danger": "#DC2626",    # 危険（赤・控えめ）

    # ステータス
    "success": "#10B981",
    "success_light": "#D1FAE5",
    "warning": "#F59E0B",
    "warning_light": "#FEF3C7",
    "error": "#EF4444",
    "error_light": "#FEE2E2",
    "info": "#3B82F6",
    "info_light": "#DBEAFE",

    # ボーダー・区切り
    "border_light": "#E5E7EB",     # 薄いボーダー
    "border_default": "#D1D5DB",   # 標準ボーダー
    "border_dark": "#9CA3AF",      # 濃いボーダー
    "divider": "#F3F4F6",          # セクション区切り

    # 差分表示
    "diff_changed": "#FEF3C7",
    "diff_added": "#D1FAE5",
    "diff_removed": "#FEE2E2",
}

# フォント設定（日本語対応）
FONT_FAMILY_SANS = "Meiryo UI"       # クリーンな日本語フォント
FONT_FAMILY_MONO = "MS Gothic"       # 日本語対応等幅フォント

def get_fonts(size_preset: str = 'medium') -> dict:
    base = {'small': 10, 'medium': 11, 'large': 13}.get(size_preset, 11)
    return {
        # 見出し系（Semibold）
        "display": (FONT_FAMILY_SANS, base + 8, "bold"),      # アプリタイトル
        "title": (FONT_FAMILY_SANS, base + 4, "bold"),        # 画面タイトル
        "heading": (FONT_FAMILY_SANS, base + 2, "bold"),      # セクション見出し

        # 本文系
        "body": (FONT_FAMILY_SANS, base, "normal"),           # 本文
        "body_medium": (FONT_FAMILY_SANS, base, "bold"),      # 本文（強調）
        "body_bold": (FONT_FAMILY_SANS, base, "bold"),        # ボタンラベル

        # 補助系
        "caption": (FONT_FAMILY_SANS, base - 1, "normal"),    # キャプション
        "small": (FONT_FAMILY_SANS, base - 2, "normal"),      # 注釈
        "tiny": (FONT_FAMILY_SANS, base - 3, "normal"),       # 極小

        # ログ・データ表示用（日本語対応）
        "mono": (FONT_FAMILY_SANS, base, "normal"),
        "mono_small": (FONT_FAMILY_SANS, base - 1, "normal"),
    }

FONTS = get_fonts('medium')

# スペーシングシステム（8pxベース）
SPACING = {
    "none": 0,
    "xs": 4,
    "sm": 8,
    "md": 12,
    "lg": 16,
    "xl": 24,
    "2xl": 32,
    "3xl": 48,
}

# 角丸
RADIUS = {
    "none": 0,
    "sm": 4,
    "default": 6,
    "md": 8,
    "lg": 12,
    "full": 9999,
}


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
        ttk.Label(toolbar, text=t('filter_label')).pack(side="left", padx=(0, 5))
        self.filter_var = tk.StringVar()
        self.filter_entry = ttk.Entry(toolbar, textvariable=self.filter_var, width=20)
        self.filter_entry.pack(side="left", padx=(0, 5))
        self.filter_var.trace_add("write", lambda *args: self._apply_filter())

        ttk.Button(toolbar, text=t('btn_clear_grid'), command=self._clear_filter).pack(side="left")

        # スペーサー
        ttk.Frame(toolbar).pack(side="left", fill="x", expand=True)

        # 一括置換ボタン
        ttk.Button(toolbar, text=t('btn_replace_all'), command=self._show_replace_dialog).pack(side="left", padx=2)

        # Undo/Redo
        self.undo_btn = ttk.Button(toolbar, text=t('btn_undo'), command=self._do_undo)
        self.undo_btn.pack(side="left", padx=2)
        self.redo_btn = ttk.Button(toolbar, text=t('btn_redo'), command=self._do_redo)
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

        self.tree.column("slide", width=80, minwidth=60, anchor="center", stretch=False)
        self.tree.column("id", width=100, minwidth=80, stretch=False)
        self.tree.column("type", width=80, minwidth=60, stretch=False)
        self.tree.column("text", width=800, minwidth=300, stretch=True)

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
        dialog.title(t('btn_replace_all'))
        dialog.geometry("400x150")
        dialog.transient(self)
        dialog.grab_set()

        ttk.Label(dialog, text=t('replace_search')).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        find_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=find_var, width=40).grid(row=0, column=1, padx=10, pady=10)

        ttk.Label(dialog, text=t('replace_with')).grid(row=1, column=0, padx=10, pady=10, sticky="w")
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
            messagebox.showinfo(t('dialog_complete'), t('result_replaced', count))

        ttk.Button(dialog, text=t('btn_replace'), command=do_replace).grid(row=2, column=1, pady=10, sticky="e")

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

        ttk.Label(frame, text=t('compare_title'), font=FONTS["heading"]).pack(anchor='w', pady=(0, 15))

        # ファイル1
        f1 = ttk.Frame(frame)
        f1.pack(fill='x', pady=5)
        ttk.Label(f1, text=t('compare_file1'), width=12).pack(side='left')
        self.file1_var = tk.StringVar()
        ttk.Entry(f1, textvariable=self.file1_var, width=45).pack(side='left', padx=5)
        ttk.Button(f1, text=t('btn_browse'), command=lambda: self._browse(self.file1_var)).pack(side='left')

        # ファイル2
        f2 = ttk.Frame(frame)
        f2.pack(fill='x', pady=5)
        ttk.Label(f2, text=t('compare_file2'), width=12).pack(side='left')
        self.file2_var = tk.StringVar()
        ttk.Entry(f2, textvariable=self.file2_var, width=45).pack(side='left', padx=5)
        ttk.Button(f2, text=t('btn_browse'), command=lambda: self._browse(self.file2_var)).pack(side='left')

        # オプション
        opt = ttk.Frame(frame)
        opt.pack(fill='x', pady=15)
        self.ignore_ws = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt, text=t('compare_ignore_ws'), variable=self.ignore_ws).pack(side='left')

        # ボタン
        btn = ttk.Frame(frame)
        btn.pack(fill='x', pady=10)
        ttk.Button(btn, text=t('btn_cancel'), command=self.dialog.destroy).pack(side='left')
        tk.Button(btn, text=t('btn_run_compare'), bg=COLOR_PALETTE["brand_primary"], fg="#FFFFFF",
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
        ttk.Label(top, text=f"📊 {stats['same']} | {stats['changed']} | {stats['added']} | {stats['removed']}",
                  font=FONTS["heading"]).pack(side='left')

        ttk.Button(top, text=t('btn_export_csv'), command=self._export_csv).pack(side='right')

        # グリッド
        grid_frame = ttk.Frame(self.window, padding=10)
        grid_frame.pack(fill='both', expand=True)

        cols = ("select", "slide", "id", "status", "before", "after")
        self.tree = ttk.Treeview(grid_frame, columns=cols, show="headings")

        self.tree.heading("select", text=t('header_select'))
        self.tree.heading("slide", text=t('header_slide'))
        self.tree.heading("id", text="ID")
        self.tree.heading("status", text=t('header_status'))
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
        ttk.Button(bottom, text=t('btn_select_original'), command=lambda: self._select_all("before")).pack(side='left', padx=2)
        ttk.Button(bottom, text=t('btn_select_new'), command=lambda: self._select_all("after")).pack(side='left', padx=2)
        tk.Button(bottom, text=t('btn_apply_selection'), font=(FONT_FAMILY_SANS, 10),
                  bg=COLOR_PALETTE["action_update"], fg="#FFFFFF", relief="flat",
                  activebackground="#047857", padx=SPACING["lg"], pady=SPACING["sm"],
                  cursor="hand2", command=self._apply).pack(side='right', padx=5)
        ttk.Button(bottom, text=t('btn_close'), command=self.window.destroy).pack(side='right')

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
            messagebox.showinfo(t('dialog_complete'), t('result_applied', len(selected)))
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
        messagebox.showinfo(t('dialog_complete'), t('result_csv_saved'))


# ============== メインアプリケーション ==============
class InsightSlidesApp:
    def __init__(self, root):
        self.root = root
        self.license_manager = LicenseManager()
        self.config_manager = ConfigManager()
        self.processing = False
        self.cancel_requested = False
        self.presentation = None
        self.log_buffer = []
        self.extracted_data = []  # グリッド用
        self.include_notes_var = tk.BooleanVar(value=False)
        self.auto_backup_var = tk.BooleanVar(value=self.config_manager.get('auto_backup', True))

        self._setup_window()
        self._apply_styles()
        self._create_menu()
        self._create_layout()
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)

        # 起動時ライセンスチェック（UIが表示された後に実行）
        self.root.after(100, self._check_license_on_startup)

    def _setup_window(self):
        tier = self.license_manager.get_tier_info()
        self.root.title(f"{APP_NAME} v{APP_VERSION} - {tier['name']}")
        self.root.geometry("1300x900")
        self.root.minsize(1100, 700)
        self.root.configure(bg=COLOR_PALETTE["bg_primary"])

    def _apply_styles(self):
        """シンプルで統一感のあるスタイル"""
        self.style = ttk.Style()
        self.style.theme_use('clam')

        # 統一背景色（全体で一貫性を持たせる）
        BG = COLOR_PALETTE["bg_primary"]  # #FFFFFF
        BG_LIGHT = COLOR_PALETTE["bg_secondary"]  # #F8FAFC
        TEXT = COLOR_PALETTE["text_primary"]  # #1F2937
        TEXT_SUB = COLOR_PALETTE["text_tertiary"]  # #6B7280
        BORDER = COLOR_PALETTE["border_light"]  # #E5E7EB

        # === フレーム ===
        self.style.configure('Main.TFrame', background=BG)
        self.style.configure('Card.TFrame', background=BG)
        self.style.configure('Sidebar.TFrame', background=BG)
        self.style.configure('TFrame', background=BG)

        # === ラベルフレーム ===
        self.style.configure('TLabelframe', background=BG, bordercolor=BORDER)
        self.style.configure('TLabelframe.Label', background=BG, foreground=TEXT,
                            font=(FONT_FAMILY_SANS, 11, "bold"))

        # === ラベル（全て同じ背景） ===
        self.style.configure('TLabel', background=BG, foreground=TEXT,
                            font=(FONT_FAMILY_SANS, 10))
        self.style.configure('Muted.TLabel', background=BG, foreground=TEXT_SUB,
                            font=(FONT_FAMILY_SANS, 9))
        self.style.configure('Caption.TLabel', background=BG, foreground=TEXT_SUB,
                            font=(FONT_FAMILY_SANS, 9))

        # === ボタン ===
        self.style.configure('TButton', background=BG_LIGHT, foreground=TEXT,
                            bordercolor=BORDER, padding=(12, 6),
                            font=(FONT_FAMILY_SANS, 10))
        self.style.map('TButton',
                      background=[('active', COLOR_PALETTE["bg_elevated"])])

        # === チェックボックス ===
        self.style.configure('TCheckbutton', background=BG, foreground=TEXT,
                            font=(FONT_FAMILY_SANS, 10))
        self.style.map('TCheckbutton', background=[('active', BG)])

        # === コンボボックス ===
        self.style.configure('TCombobox', fieldbackground=BG, background=BG,
                            foreground=TEXT, bordercolor=BORDER,
                            padding=(4, 2), font=(FONT_FAMILY_SANS, 10))

        # === エントリ ===
        self.style.configure('TEntry', fieldbackground=BG, foreground=TEXT,
                            bordercolor=BORDER, padding=(4, 2))

        # === Notebook（タブ） ===
        self.style.configure('TNotebook', background=BG, bordercolor=BORDER)
        self.style.configure('TNotebook.Tab', background=BG_LIGHT, foreground=TEXT_SUB,
                            padding=(16, 8), font=(FONT_FAMILY_SANS, 10))
        self.style.map('TNotebook.Tab',
                      background=[('selected', BG)],
                      foreground=[('selected', TEXT)])

        # === プログレスバー ===
        self.style.configure('TProgressbar', background=COLOR_PALETTE["brand_primary"],
                            troughcolor=BG_LIGHT, bordercolor=BORDER)

        # === Treeview ===
        self.style.configure('Treeview', background=BG, foreground=TEXT,
                            fieldbackground=BG, bordercolor=BORDER,
                            rowheight=28, font=(FONT_FAMILY_SANS, 10))
        self.style.configure('Treeview.Heading', background=BG_LIGHT, foreground=TEXT,
                            bordercolor=BORDER, font=(FONT_FAMILY_SANS, 10, "bold"))
        self.style.map('Treeview',
                      background=[('selected', COLOR_PALETTE["brand_light"])],
                      foreground=[('selected', COLOR_PALETTE["brand_primary"])])

        # === スクロールバー ===
        self.style.configure('Vertical.TScrollbar', background=BG_LIGHT,
                            troughcolor=BG, bordercolor=BG)
        self.style.configure('Horizontal.TScrollbar', background=BG_LIGHT,
                            troughcolor=BG, bordercolor=BG)

    def _create_menu(self):
        # メニュースタイル設定
        menu_font = (FONT_FAMILY_SANS, 9)
        menu_style = {
            'font': menu_font,
            'bg': COLOR_PALETTE["bg_primary"],
            'fg': COLOR_PALETTE["text_primary"],
            'activebackground': COLOR_PALETTE["brand_primary"],
            'activeforeground': '#FFFFFF',
            'relief': 'flat',
            'bd': 0,
        }

        menubar = tk.Menu(self.root, **menu_style)
        self.root.config(menu=menubar)

        help_menu = tk.Menu(menubar, tearoff=0, **menu_style)
        menubar.add_cascade(label=t('menu_help'), menu=help_menu)
        help_menu.add_command(label=t('menu_guide'), command=lambda: webbrowser.open(SUPPORT_LINKS["tutorial"]))
        help_menu.add_command(label=t('menu_faq'), command=lambda: webbrowser.open(SUPPORT_LINKS["faq"]))
        help_menu.add_separator()
        help_menu.add_command(label=t('menu_license'), command=self._show_license_dialog)
        help_menu.add_separator()

        lang_menu = tk.Menu(help_menu, tearoff=0, **menu_style)
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
        """洗練されたヘッダー - 余白で区切り、シンプルに"""
        header = tk.Frame(parent, bg=COLOR_PALETTE["bg_primary"])
        header.grid(row=0, column=0, sticky='ew', pady=(0, SPACING["lg"]))

        # 左: タイトル + サブタイトル
        left = tk.Frame(header, bg=COLOR_PALETTE["bg_primary"])
        left.pack(side='left')

        # アプリ名（シンプルに）
        tk.Label(left, text="Insight Slides", font=FONTS["display"],
                 fg=COLOR_PALETTE["text_primary"], bg=COLOR_PALETTE["bg_primary"]).pack(side='left')

        # ライセンスバッジ（あれば）
        tier = self.license_manager.get_tier_info()
        if tier['name'] != 'Free':
            badge = tk.Label(left, text=f" {tier['name']} ", font=FONTS["small"],
                            fg=COLOR_PALETTE["brand_primary"], bg=COLOR_PALETTE["brand_light"],
                            padx=6, pady=2)
            badge.pack(side='left', padx=(SPACING["sm"], 0))

        # サブタイトル
        tk.Label(left, text=t('app_subtitle'), font=FONTS["caption"],
                 fg=COLOR_PALETTE["text_muted"], bg=COLOR_PALETTE["bg_primary"]).pack(side='left', padx=(SPACING["lg"], 0))

        # 右: バージョン
        right = tk.Frame(header, bg=COLOR_PALETTE["bg_primary"])
        right.pack(side='right')

        tk.Label(right, text=f"v{APP_VERSION}", font=FONTS["small"],
                 fg=COLOR_PALETTE["text_muted"], bg=COLOR_PALETTE["bg_primary"]).pack(side='right')

    def _create_controls(self, parent):
        """左サイドバー - 3セクション構成（入力/出力/フォルダ一括）"""
        frame = ttk.Frame(parent, style='Sidebar.TFrame')
        frame.grid(row=0, column=0, sticky='nsew', padx=(0, SPACING["xl"]))
        frame.grid_rowconfigure(4, weight=1)

        btn_font = (FONT_FAMILY_SANS, 10)
        can_json = self.license_manager.can_json()
        can_batch = self.license_manager.can_batch()
        can_compare = self.license_manager.can_compare()

        # ============ 入力（1ファイル）セクション ============
        input_card = ttk.LabelFrame(frame, text=t('panel_input'), padding=SPACING["md"])
        input_card.grid(row=0, column=0, sticky='ew', pady=(0, SPACING["md"]))
        input_card.grid_columnconfigure(0, weight=1)

        # PPTX読込ボタン（プライマリ）
        tk.Button(input_card, text=t('btn_load_pptx'), font=btn_font,
                  bg=COLOR_PALETTE["brand_primary"], fg="#FFFFFF", relief="flat",
                  activebackground=COLOR_PALETTE["brand_hover"],
                  padx=SPACING["lg"], pady=SPACING["sm"],
                  cursor="hand2", command=self._extract_single).grid(row=0, column=0, sticky='ew', pady=(0, SPACING["xs"]))

        # Excel読込ボタン
        tk.Button(input_card, text=t('btn_load_excel'), font=btn_font,
                  bg=COLOR_PALETTE["secondary_default"], fg=COLOR_PALETTE["text_secondary"], relief="flat",
                  activebackground=COLOR_PALETTE["secondary_hover"],
                  padx=SPACING["md"], pady=SPACING["sm"],
                  cursor="hand2", command=self._load_excel_to_grid).grid(row=1, column=0, sticky='ew', pady=(0, SPACING["xs"]))

        # JSON読込ボタン（Pro）
        if can_json:
            tk.Button(input_card, text=t('btn_load_json'), font=btn_font,
                      bg=COLOR_PALETTE["secondary_default"], fg=COLOR_PALETTE["text_secondary"], relief="flat",
                      activebackground=COLOR_PALETTE["secondary_hover"],
                      padx=SPACING["md"], pady=SPACING["sm"],
                      cursor="hand2", command=self._load_json_to_grid).grid(row=2, column=0, sticky='ew')
        else:
            tk.Label(input_card, text=f"{t('btn_load_json')} (Pro)", font=btn_font,
                     fg=COLOR_PALETTE["text_muted"], bg=COLOR_PALETTE["bg_primary"]).grid(row=2, column=0, sticky='w')

        # ============ 出力（1ファイル）セクション ============
        output_card = ttk.LabelFrame(frame, text=t('panel_output_file'), padding=SPACING["md"])
        output_card.grid(row=1, column=0, sticky='ew', pady=(0, SPACING["md"]))
        output_card.grid_columnconfigure(0, weight=1)

        # スライド制限警告
        limit = self.license_manager.get_update_limit()
        if limit:
            warn_frame = tk.Frame(output_card, bg=COLOR_PALETTE["warning_light"], padx=SPACING["sm"], pady=SPACING["xs"])
            warn_frame.grid(row=0, column=0, sticky='ew', pady=(0, SPACING["sm"]))
            tk.Label(warn_frame, text=t('msg_update_limit', limit), font=FONTS["small"],
                    fg=COLOR_PALETTE["warning"], bg=COLOR_PALETTE["warning_light"]).pack(anchor='w')

        # PPTXに反映ボタン（プライマリ）
        tk.Button(output_card, text=t('btn_apply_pptx'), font=btn_font,
                  bg=COLOR_PALETTE["action_update"], fg="#FFFFFF", relief="flat",
                  activebackground="#047857",
                  padx=SPACING["lg"], pady=SPACING["sm"],
                  cursor="hand2", command=self._apply_grid_to_pptx).grid(row=1, column=0, sticky='ew', pady=(0, SPACING["xs"]))

        # Excel出力ボタン
        tk.Button(output_card, text=t('btn_export_to_excel'), font=btn_font,
                  bg=COLOR_PALETTE["secondary_default"], fg=COLOR_PALETTE["text_secondary"], relief="flat",
                  activebackground=COLOR_PALETTE["secondary_hover"],
                  padx=SPACING["md"], pady=SPACING["sm"],
                  cursor="hand2", command=self._export_grid_excel).grid(row=2, column=0, sticky='ew', pady=(0, SPACING["xs"]))

        # JSON出力ボタン（Pro）
        if can_json:
            tk.Button(output_card, text=t('btn_export_to_json'), font=btn_font,
                      bg=COLOR_PALETTE["secondary_default"], fg=COLOR_PALETTE["text_secondary"], relief="flat",
                      activebackground=COLOR_PALETTE["secondary_hover"],
                      padx=SPACING["md"], pady=SPACING["sm"],
                      cursor="hand2", command=self._export_grid_json).grid(row=3, column=0, sticky='ew')
        else:
            tk.Label(output_card, text=f"{t('btn_export_to_json')} (Pro)", font=btn_font,
                     fg=COLOR_PALETTE["text_muted"], bg=COLOR_PALETTE["bg_primary"]).grid(row=3, column=0, sticky='w')

        # ============ フォルダ一括セクション ============
        batch_card = ttk.LabelFrame(frame, text=t('panel_batch'), padding=SPACING["md"])
        batch_card.grid(row=2, column=0, sticky='ew', pady=(0, SPACING["md"]))
        batch_card.grid_columnconfigure(0, weight=1)

        if can_batch:
            # Proバッジ
            tk.Label(batch_card, text="Pro", font=(FONT_FAMILY_SANS, 8, 'bold'),
                     fg=COLOR_PALETTE["brand_primary"], bg=COLOR_PALETTE["bg_primary"]).grid(row=0, column=0, sticky='e')

            # 一括抽出ボタン
            tk.Button(batch_card, text=t('btn_batch_extract'), font=btn_font,
                      bg=COLOR_PALETTE["secondary_default"], fg=COLOR_PALETTE["text_secondary"], relief="flat",
                      activebackground=COLOR_PALETTE["secondary_hover"],
                      padx=SPACING["md"], pady=SPACING["sm"],
                      cursor="hand2", command=self._batch_extract_dialog).grid(row=1, column=0, sticky='ew', pady=(0, SPACING["xs"]))

            # 一括更新ボタン
            tk.Button(batch_card, text=t('btn_batch_update'), font=btn_font,
                      bg=COLOR_PALETTE["secondary_default"], fg=COLOR_PALETTE["text_secondary"], relief="flat",
                      activebackground=COLOR_PALETTE["secondary_hover"],
                      padx=SPACING["md"], pady=SPACING["sm"],
                      cursor="hand2", command=self._batch_update_dialog).grid(row=2, column=0, sticky='ew')
        else:
            tk.Label(batch_card, text=f"{t('btn_batch_extract')} (Pro)", font=btn_font,
                     fg=COLOR_PALETTE["text_muted"], bg=COLOR_PALETTE["bg_primary"]).grid(row=0, column=0, sticky='w', pady=(0, SPACING["xs"]))
            tk.Label(batch_card, text=f"{t('btn_batch_update')} (Pro)", font=btn_font,
                     fg=COLOR_PALETTE["text_muted"], bg=COLOR_PALETTE["bg_primary"]).grid(row=1, column=0, sticky='w')

        # ============ 2ファイル比較ボタン ============
        compare_text = t('btn_compare') if can_compare else f"{t('btn_compare')} (STD)"
        tk.Button(frame, text=compare_text, font=btn_font,
                  bg=COLOR_PALETTE["secondary_default"] if can_compare else COLOR_PALETTE["bg_secondary"],
                  fg=COLOR_PALETTE["text_secondary"] if can_compare else COLOR_PALETTE["text_muted"],
                  activebackground=COLOR_PALETTE["secondary_hover"] if can_compare else COLOR_PALETTE["bg_secondary"],
                  relief="flat", padx=SPACING["md"], pady=SPACING["sm"],
                  cursor="hand2" if can_compare else "arrow",
                  command=self._show_compare_dialog if can_compare else None,
                  state='normal' if can_compare else 'disabled').grid(row=3, column=0, sticky='ew', pady=(0, SPACING["md"]))

        # ステータス＆ミニログ
        status_frame = ttk.Frame(frame, style='Main.TFrame')
        status_frame.grid(row=4, column=0, sticky='sew')

        # プログレスバー（処理中のみ表示）
        self.progress = ttk.Progressbar(status_frame, mode='indeterminate')
        self.progress.pack(fill='x', pady=(0, SPACING["sm"]))

        # ミニログ（1-2行、クリックで詳細表示）
        log_frame = tk.Frame(status_frame, bg=COLOR_PALETTE["bg_secondary"], cursor="hand2")
        log_frame.pack(fill='x')
        log_frame.bind("<Button-1>", lambda e: self._show_log_detail())

        self.mini_log_label = tk.Label(log_frame, text=t('status_waiting'),
                                       font=(FONT_FAMILY_SANS, 9), fg=COLOR_PALETTE["text_tertiary"],
                                       bg=COLOR_PALETTE["bg_secondary"], anchor='w', padx=SPACING["sm"], pady=SPACING["xs"])
        self.mini_log_label.pack(fill='x')
        self.mini_log_label.bind("<Button-1>", lambda e: self._show_log_detail())

        # 詳細リンク
        detail_link = tk.Label(log_frame, text=t('show_detail'),
                               font=(FONT_FAMILY_SANS, 8), fg=COLOR_PALETTE["brand_primary"],
                               bg=COLOR_PALETTE["bg_secondary"], cursor="hand2")
        detail_link.pack(anchor='e', padx=SPACING["sm"], pady=(0, SPACING["xs"]))
        detail_link.bind("<Button-1>", lambda e: self._show_log_detail())

        # キャンセルボタン（処理中のみアクティブ）
        btn_frame = ttk.Frame(status_frame)
        btn_frame.pack(fill='x', pady=(SPACING["sm"], 0))
        self.cancel_btn = ttk.Button(btn_frame, text=t('btn_cancel'), command=self._cancel, state='disabled')
        self.cancel_btn.pack(side='left')

    def _create_output(self, parent):
        """右側メインコンテンツ - 編集専用エリア"""
        # メインカード（タイトルなし - 構造で役割を示す）
        card = ttk.Frame(parent, style='Card.TFrame')
        card.grid(row=0, column=1, sticky='nsew')
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(1, weight=1)

        # ファイル情報ヘッダー（コンパクト）
        file_info_frame = tk.Frame(card, bg=COLOR_PALETTE["bg_primary"], pady=SPACING["sm"])
        file_info_frame.grid(row=0, column=0, sticky='ew', padx=SPACING["md"])

        self.file_name_label = tk.Label(file_info_frame, text="",
                                        font=(FONT_FAMILY_SANS, 10), bg=COLOR_PALETTE["bg_primary"],
                                        fg=COLOR_PALETTE["text_secondary"])
        self.file_name_label.pack(side='left')

        self.file_info_detail = tk.Label(file_info_frame, text="",
                                         font=FONTS["caption"], bg=COLOR_PALETTE["bg_primary"],
                                         fg=COLOR_PALETTE["text_tertiary"])
        self.file_info_detail.pack(side='right')

        # メイン編集エリア（グリッド）
        edit_area = ttk.Frame(card, style='Main.TFrame')
        edit_area.grid(row=1, column=0, sticky='nsew', padx=SPACING["md"])
        edit_area.grid_columnconfigure(0, weight=1)
        edit_area.grid_rowconfigure(0, weight=1)

        # ウェルカムガイド（初期状態）
        self.welcome_frame = tk.Frame(edit_area, bg=COLOR_PALETTE["bg_primary"])
        self.welcome_frame.grid(row=0, column=0, sticky='nsew')
        self.welcome_frame.grid_columnconfigure(0, weight=1)
        self.welcome_frame.grid_rowconfigure(0, weight=1)
        self._create_welcome_guide()

        # グリッドビュー（データ読込後に表示）
        self.grid_container = ttk.Frame(edit_area, style='Main.TFrame')
        self.grid_view = EditableGrid(self.grid_container, on_change=self._on_grid_change)
        self.grid_view.grid(row=0, column=0, sticky='nsew')
        self.grid_container.grid_columnconfigure(0, weight=1)
        self.grid_container.grid_rowconfigure(0, weight=1)

        # アクションバー（下部固定）
        action_bar = tk.Frame(card, bg=COLOR_PALETTE["bg_primary"], pady=SPACING["md"])
        action_bar.grid(row=2, column=0, sticky='ew', padx=SPACING["md"])

        # プライマリアクション
        self.apply_btn = tk.Button(action_bar, text=t('btn_apply'), font=(FONT_FAMILY_SANS, 10),
                  bg=COLOR_PALETTE["action_update"], fg="#FFFFFF", relief="flat",
                  padx=SPACING["lg"], pady=SPACING["sm"],
                  activebackground="#047857",
                  cursor="hand2", command=self._apply_grid_to_pptx, state='disabled')
        self.apply_btn.pack(side='right')

        # エクスポートボタン（Excel）
        self.export_excel_btn = tk.Button(action_bar, text=t('btn_export_excel'), font=(FONT_FAMILY_SANS, 10),
                  bg=COLOR_PALETTE["secondary_default"], fg=COLOR_PALETTE["text_secondary"], relief="flat",
                  padx=SPACING["md"], pady=SPACING["sm"],
                  activebackground=COLOR_PALETTE["secondary_hover"],
                  cursor="hand2", command=self._export_grid_excel, state='disabled')
        self.export_excel_btn.pack(side='right', padx=(0, SPACING["sm"]))

        # エクスポートボタン（JSON）- Pro版以上
        if self.license_manager.can_json():
            self.export_json_btn = tk.Button(action_bar, text=t('btn_export_json'), font=(FONT_FAMILY_SANS, 10),
                      bg=COLOR_PALETTE["secondary_default"], fg=COLOR_PALETTE["text_secondary"], relief="flat",
                      padx=SPACING["md"], pady=SPACING["sm"],
                      activebackground=COLOR_PALETTE["secondary_hover"],
                      cursor="hand2", command=self._export_grid_json, state='disabled')
            self.export_json_btn.pack(side='right', padx=(0, SPACING["sm"]))
        else:
            self.export_json_btn = tk.Button(action_bar, text=f"{t('btn_export_json')} (Pro)", font=(FONT_FAMILY_SANS, 10),
                      bg=COLOR_PALETTE["bg_secondary"], fg=COLOR_PALETTE["text_muted"], relief="flat",
                      padx=SPACING["md"], pady=SPACING["sm"], state='disabled')
            self.export_json_btn.pack(side='right', padx=(0, SPACING["sm"]))

    def _create_welcome_guide(self):
        """初期状態のウェルカムガイド"""
        center_frame = tk.Frame(self.welcome_frame, bg=COLOR_PALETTE["bg_primary"])
        center_frame.place(relx=0.5, rely=0.45, anchor='center')

        # タイトル
        tk.Label(center_frame, text=t('welcome_guide_title'),
                 font=(FONT_FAMILY_SANS, 16, "bold"), fg=COLOR_PALETTE["text_primary"],
                 bg=COLOR_PALETTE["bg_primary"]).pack(pady=(0, SPACING["lg"]))

        # 手順
        steps = [
            ("1", t('guide_step1')),
            ("2", t('guide_step2')),
            ("3", t('guide_step3')),
            ("4", t('guide_step4')),
        ]

        for num, text in steps:
            step_frame = tk.Frame(center_frame, bg=COLOR_PALETTE["bg_primary"])
            step_frame.pack(anchor='w', pady=SPACING["xs"])

            tk.Label(step_frame, text=num, font=(FONT_FAMILY_SANS, 10, "bold"),
                     fg=COLOR_PALETTE["brand_primary"], bg=COLOR_PALETTE["bg_primary"],
                     width=2).pack(side='left')
            tk.Label(step_frame, text=text, font=(FONT_FAMILY_SANS, 10),
                     fg=COLOR_PALETTE["text_secondary"], bg=COLOR_PALETTE["bg_primary"]).pack(side='left')

    def _show_edit_area(self):
        """ウェルカムガイドを隠してグリッドを表示"""
        self.welcome_frame.grid_remove()
        self.grid_container.grid(row=0, column=0, sticky='nsew')
        self.apply_btn.configure(state='normal')
        self.export_excel_btn.configure(state='normal')
        self.export_json_btn.configure(state='normal')

    def _show_welcome_area(self):
        """グリッドを隠してウェルカムガイドを表示"""
        self.grid_container.grid_remove()
        self.welcome_frame.grid(row=0, column=0, sticky='nsew')
        self.apply_btn.configure(state='disabled')
        self.export_excel_btn.configure(state='disabled')
        self.export_json_btn.configure(state='disabled')
        self.file_name_label.configure(text="")
        self.file_info_detail.configure(text="")

    def _update_file_info(self, filename: str, item_count: int = 0, slide_count: int = 0):
        """ファイル情報ヘッダーを更新"""
        self.file_name_label.configure(text=filename)
        if item_count > 0:
            self.file_info_detail.configure(text=t('status_slides_items', slide_count, item_count))
        else:
            self.file_info_detail.configure(text="")

    def _show_welcome(self):
        """初期ウェルカム表示（ミニログのみ更新）"""
        tier = self.license_manager.get_tier_info()
        self._update_mini_log(f"{APP_NAME} v{APP_VERSION} ({tier['name']}) - 準備完了")

    # === Output helpers ===
    def _update_output(self, text, clear=False):
        """ログバッファに追加し、ミニログを更新"""
        if clear:
            self.log_buffer = []
        self.log_buffer.append(text)
        # ミニログには最新の1行のみ表示
        last_line = text.strip().split('\n')[-1] if text.strip() else ""
        self._update_mini_log(last_line)

    def _update_output_safe(self, text, clear=False):
        self.root.after(0, lambda: self._update_output(text, clear))

    def _update_mini_log(self, text):
        """ミニログラベルを更新（最新メッセージのみ）"""
        # 長すぎるテキストは省略
        max_len = 50
        display_text = text[:max_len] + "..." if len(text) > max_len else text
        self.mini_log_label.configure(text=display_text)

    def _update_mini_log_safe(self, text):
        self.root.after(0, lambda: self._update_mini_log(text))

    def _update_status(self, text, color=None):
        """ステータス更新（ミニログに統合）"""
        self._update_mini_log(text)

    def _update_status_safe(self, text, color=None):
        self.root.after(0, lambda: self._update_status(text, color))

    def _log(self, msg, level="info"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        prefix = {"error": "× ", "warning": "! ", "success": "✓ "}.get(level, "")
        full_msg = f"[{timestamp}] {prefix}{msg}"
        self._update_output_safe(f"{full_msg}\n")
        # エラー時は色を変える
        if level == "error":
            self.mini_log_label.configure(fg=COLOR_PALETTE["error"])
        elif level == "success":
            self.mini_log_label.configure(fg=COLOR_PALETTE["success"])
        else:
            self.mini_log_label.configure(fg=COLOR_PALETTE["text_tertiary"])

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
            self._log(t('log_cancel_request'), "warning")

    def _show_log_detail(self):
        """ログ詳細モーダルを表示"""
        dialog = tk.Toplevel(self.root)
        dialog.title("処理ログ")
        dialog.geometry("600x400")
        dialog.transient(self.root)
        dialog.grab_set()

        # ログテキストエリア
        text_frame = ttk.Frame(dialog)
        text_frame.pack(fill='both', expand=True, padx=SPACING["md"], pady=SPACING["md"])

        log_text = scrolledtext.ScrolledText(text_frame, wrap=tk.WORD,
                                              font=FONTS["mono"],
                                              bg=COLOR_PALETTE["bg_primary"],
                                              fg=COLOR_PALETTE["text_secondary"],
                                              relief="flat", bd=1)
        log_text.pack(fill='both', expand=True)

        # ログ内容を表示
        log_content = "".join(self.log_buffer) if self.log_buffer else "ログはありません"
        log_text.insert('1.0', log_content)
        log_text.configure(state=tk.DISABLED)
        log_text.see(tk.END)

        # ボタンフレーム
        btn_frame = tk.Frame(dialog, bg=COLOR_PALETTE["bg_primary"])
        btn_frame.pack(fill='x', padx=SPACING["md"], pady=(0, SPACING["md"]))

        def copy_log():
            content = "".join(self.log_buffer)
            if content:
                self.root.clipboard_clear()
                self.root.clipboard_append(content)
                messagebox.showinfo("コピー完了", "ログをクリップボードにコピーしました")

        def clear_log():
            self.log_buffer = []
            log_text.configure(state=tk.NORMAL)
            log_text.delete('1.0', tk.END)
            log_text.insert('1.0', "ログをクリアしました")
            log_text.configure(state=tk.DISABLED)
            self._update_mini_log("準備完了")

        tk.Button(btn_frame, text=t('btn_copy_log'), font=(FONT_FAMILY_SANS, 9),
                  bg=COLOR_PALETTE["secondary_default"], fg=COLOR_PALETTE["text_secondary"],
                  relief="flat", padx=SPACING["md"], pady=SPACING["xs"],
                  command=copy_log).pack(side='left', padx=(0, SPACING["sm"]))

        tk.Button(btn_frame, text=t('btn_clear_log'), font=(FONT_FAMILY_SANS, 9),
                  bg=COLOR_PALETTE["secondary_default"], fg=COLOR_PALETTE["text_secondary"],
                  relief="flat", padx=SPACING["md"], pady=SPACING["xs"],
                  command=clear_log).pack(side='left')

        tk.Button(btn_frame, text=t('btn_close'), font=(FONT_FAMILY_SANS, 9),
                  bg=COLOR_PALETTE["brand_primary"], fg="#FFFFFF",
                  relief="flat", padx=SPACING["md"], pady=SPACING["xs"],
                  command=dialog.destroy).pack(side='right')

    def _change_language(self, lang):
        if lang != get_language():
            self.config_manager.set('language', lang)
            set_language(lang)
            self._create_menu()  # メニューを再作成
            self._create_layout()
            self._setup_window()  # ウィンドウタイトルを更新
            messagebox.showinfo(t('dialog_complete'), t('lang_changed'))

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
                self._update_status_safe(t('log_processing'))
                self._update_output_safe(f"\n📄 処理開始: {os.path.basename(path)}\n", clear=True)

                data, meta = self.extract_from_ppt(path, include_notes)
                if self.cancel_requested:
                    return self._log(t('log_cancelled'), "warning")

                if data:
                    # ファイル情報を更新
                    filename = os.path.basename(path)
                    slide_count = meta.get('slide_count', 0)
                    self.root.after(0, lambda: self._update_file_info(filename, len(data), slide_count))

                    # グリッドにロード
                    self.extracted_data = data
                    self.root.after(0, lambda: self.grid_view.load_data(data))
                    self.root.after(0, lambda: self._show_edit_area())

                    # ファイル保存（デフォルトはExcel）
                    out = os.path.splitext(path)[0] + "_抽出.xlsx"
                    if self.save_to_file(data, out, "excel"):
                        self._log(f"✅ {t('status_complete_items', len(data))} → {os.path.basename(out)}", "success")
                        self._update_status_safe(t('status_complete_items', len(data)))
                else:
                    self._log(t('log_no_text'), "warning")
            except Exception as e:
                save_error_log(e, "_extract_single")
                self._log(t('log_error', e), "error")
            finally:
                self._stop_progress()

        threading.Thread(target=run, daemon=True).start()

    def _extract_batch(self, format: str = "excel"):
        """フォルダ一括抽出 (excel/json)"""
        if self.processing:
            return
        folder = filedialog.askdirectory(title=t('dialog_select_folder'))
        if not folder:
            return

        include_notes = self.include_notes_var.get() if self.license_manager.is_pro() else False
        ext = ".xlsx" if format == "excel" else ".json"

        def run():
            try:
                self._start_progress()
                self._update_output_safe(f"\n📁 フォルダ一括出力 ({format.upper()}): {folder}\n", clear=True)

                files = [f for f in Path(folder).glob("*.pptx") if not f.name.startswith("~$")]
                if not files:
                    return self._log(t('log_no_pptx_found'), "warning")

                self._log(t('log_found_files', len(files)))
                total = 0

                for i, f in enumerate(files, 1):
                    if self.cancel_requested:
                        break
                    self._log(f"[{i}/{len(files)}] {f.name}")
                    data, meta = self.extract_from_ppt(str(f), include_notes)
                    if data:
                        out = str(f.with_suffix('')) + f"_抽出{ext}"
                        self.save_to_file(data, out, format)
                        total += len(data)

                self._log(f"✅ {t('status_batch_complete', total, format.upper())}", "success")
            except Exception as e:
                self._log(t('log_error', e), "error")
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
                    self._log(t('log_invalid_header'), "error")
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
                    return self._log(t('log_no_update_data'), "warning")

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
                        messagebox.showinfo(t('dialog_complete'), t('result_updated', updated, skipped))

                self.root.after(0, save)
            except Exception as e:
                self._log(t('log_error', e), "error")
            finally:
                self._stop_progress()

        threading.Thread(target=run, daemon=True).start()

    def _update_excel(self):
        self._run_update("excel")

    def _update_json(self):
        self._run_update("json")

    def _update_batch(self, format: str = "excel"):
        """フォルダ内のExcel/JSONファイルとPPTXを一括更新"""
        if self.processing:
            return

        ext = ".xlsx" if format == "excel" else ".json"
        folder = filedialog.askdirectory(title=t('dialog_select_folder_update', ext))
        if not folder:
            return

        def run():
            try:
                self._start_progress()
                self._update_output_safe(f"\n📁 フォルダ一括読込 ({format.upper()}): {folder}\n", clear=True)

                folder_path = Path(folder)

                # 指定形式のファイルを検索
                data_files = list(folder_path.glob(f"*_抽出{ext}"))

                if not data_files:
                    return self._log(f"抽出ファイル (*_抽出{ext}) が見つかりません", "warning")

                self._log(t('log_found_files', len(data_files)))
                updated_count = 0
                error_count = 0

                for i, data_file in enumerate(data_files, 1):
                    if self.cancel_requested:
                        break

                    # 対応するPPTXファイルを検索
                    base_name = data_file.stem.replace("_抽出", "")
                    pptx_path = folder_path / f"{base_name}.pptx"

                    if not pptx_path.exists():
                        self._log(f"[{i}/{len(data_files)}] {data_file.name}: PPTXなし (スキップ)", "warning")
                        continue

                    self._log(f"[{i}/{len(data_files)}] {pptx_path.name}")

                    try:
                        updates = self._load_updates(str(data_file), format)

                        if not updates:
                            self._log(f"  → 更新データなし", "warning")
                            continue

                        self._create_backup(str(pptx_path))
                        updated, skipped, _ = self._update_ppt(str(pptx_path), updates)

                        # 更新済みファイルを保存
                        out_path = folder_path / f"{base_name}_更新済み.pptx"
                        self.presentation.save(str(out_path))
                        self._log(f"  → {updated}件更新, 保存: {out_path.name}")
                        updated_count += 1

                    except Exception as e:
                        self._log(f"  → エラー: {e}", "error")
                        error_count += 1

                self._log(f"\n✅ バッチ読込完了 ({format.upper()}): {updated_count}件成功, {error_count}件エラー", "success")

            except Exception as e:
                self._log(t('log_error', e), "error")
            finally:
                self._stop_progress()

        threading.Thread(target=run, daemon=True).start()

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
                    return self._log(t('log_no_update_data'), "warning")

                _, _, changes = self._update_ppt(ppt_path, updates, preview=True)
                if changes:
                    self._log(f"\n変更箇所: {len(changes)}件")
                    for i, c in enumerate(changes[:20], 1):
                        self._update_output_safe(f"[{i}] スライド{c['slide']} ID:{c['id']}\n  旧: {c['old']}\n  新: {c['new']}\n\n")
                else:
                    self._log("変更箇所なし")
            except Exception as e:
                self._log(t('log_error', e), "error")
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
                self._log(t('log_error', e), "error")
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
        self._show_edit_area()

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
                        messagebox.showinfo(t('dialog_complete'), t('status_update_complete', updated))

                self.root.after(0, save)
            except Exception as e:
                self._log(t('log_error', e), "error")
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
            messagebox.showinfo(t('dialog_complete'), t('result_export_complete', path))

    def _export_grid_json(self):
        data = self.grid_view.get_data()
        if not data:
            messagebox.showwarning("警告", "エクスポートするデータがありません")
            return

        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
        if not path:
            return

        if self.save_to_file(data, path, "json"):
            messagebox.showinfo(t('dialog_complete'), t('result_export_complete', path))

    def _load_excel_to_grid(self):
        """Excelファイルをグリッドにロードする"""
        path = filedialog.askopenfilename(
            title=t('btn_load_excel'),
            filetypes=[("Excel", "*.xlsx")]
        )
        if not path:
            return

        try:
            wb = openpyxl.load_workbook(path)
            ws = wb.active
            headers = [c.value for c in ws[1]]

            # ヘッダー列を特定
            try:
                si = headers.index("スライド番号") if "スライド番号" in headers else headers.index("slide")
                oi = headers.index("オブジェクトID") if "オブジェクトID" in headers else headers.index("id")
                ti = headers.index("テキスト内容") if "テキスト内容" in headers else headers.index("text")
                type_i = headers.index("タイプ") if "タイプ" in headers else (headers.index("type") if "type" in headers else None)
            except ValueError:
                messagebox.showerror(t('dialog_error'), t('log_invalid_header'))
                return

            data = []
            for row in list(ws.rows)[1:]:
                try:
                    slide = str(row[si].value) if row[si].value else ""
                    oid = str(row[oi].value) if row[oi].value else ""
                    txt = str(row[ti].value) if row[ti].value else ""
                    obj_type = str(row[type_i].value) if type_i is not None and row[type_i].value else ""
                    if txt == "None":
                        txt = ""
                    if slide and oid:
                        data.append({
                            "slide": slide,
                            "id": oid,
                            "type": obj_type,
                            "text": txt
                        })
                except Exception:
                    pass

            if data:
                self.extracted_data = data
                self.grid_view.load_data(data)
                self._show_edit_area()
                self._update_file_info(os.path.basename(path), len(data))
                self._log(f"✅ {t('status_complete_items', len(data))}", "success")
            else:
                messagebox.showwarning(t('dialog_error'), t('log_no_text'))

        except Exception as e:
            save_error_log(e, "_load_excel_to_grid")
            messagebox.showerror(t('dialog_error'), str(e))

    def _load_json_to_grid(self):
        """JSONファイルをグリッドにロードする"""
        if not self.license_manager.can_json():
            messagebox.showinfo(t('dialog_error'), "JSON機能はPro版以上が必要です")
            return

        path = filedialog.askopenfilename(
            title=t('btn_load_json'),
            filetypes=[("JSON", "*.json")]
        )
        if not path:
            return

        try:
            with open(path, 'r', encoding='utf-8') as f:
                json_data = json.load(f)

            data = []
            for item in json_data:
                slide = str(item.get('スライド番号') or item.get('slide', ''))
                oid = str(item.get('オブジェクトID') or item.get('id', ''))
                txt = str(item.get('テキスト内容') or item.get('text', ''))
                obj_type = str(item.get('タイプ') or item.get('type', ''))
                if slide and oid:
                    data.append({
                        "slide": slide,
                        "id": oid,
                        "type": obj_type,
                        "text": txt
                    })

            if data:
                self.extracted_data = data
                self.grid_view.load_data(data)
                self._show_edit_area()
                self._update_file_info(os.path.basename(path), len(data))
                self._log(f"✅ {t('status_complete_items', len(data))}", "success")
            else:
                messagebox.showwarning(t('dialog_error'), t('log_no_text'))

        except Exception as e:
            save_error_log(e, "_load_json_to_grid")
            messagebox.showerror(t('dialog_error'), str(e))

    def _batch_extract_dialog(self):
        """フォルダ一括抽出（Excel形式）"""
        self._extract_batch("excel")

    def _batch_update_dialog(self):
        """フォルダ一括更新（Excel形式）"""
        self._update_batch("excel")

    # === Dialogs ===
    def _check_license_on_startup(self):
        """起動時のライセンスチェック"""
        # ライセンスが未アクティベートの場合、認証ダイアログを表示
        if not self.license_manager.is_activated():
            self._show_license_dialog(startup_check=True)
            return

        # 有効期限切れチェック
        days = self.license_manager.get_days_until_expiry()
        if days is not None and days <= 0:
            # 期限切れ - Free版にダウングレード
            messagebox.showwarning(
                t('dialog_error'),
                t('license_expired')
            )
            self.license_manager.deactivate()
            self._create_layout()
            self._show_license_dialog(startup_check=True)
            return

        # 期限切れ警告（30日以内）
        if self.license_manager.should_show_expiry_warning():
            expiry_str = self.license_manager.get_expiry_date_str()
            messagebox.showinfo(
                t('license_title'),
                t('license_expiry_warning', days, expiry_str)
            )

    def _show_license_dialog(self, startup_check: bool = False):
        """ライセンス認証ダイアログを表示

        Args:
            startup_check: 起動時チェックの場合True（キャンセル不可）
        """
        dialog = tk.Toplevel(self.root)
        dialog.title(t('license_auth_title'))
        dialog.geometry("550x520")
        dialog.minsize(550, 520)
        dialog.transient(self.root)
        dialog.grab_set()

        if startup_check:
            dialog.protocol("WM_DELETE_WINDOW", lambda: None)  # 閉じるボタン無効

        frame = ttk.Frame(dialog, padding=20)
        frame.pack(fill='both', expand=True)

        # ヘッダー
        header_frame = ttk.Frame(frame)
        header_frame.pack(fill='x', pady=(0, 15))
        ttk.Label(header_frame, text=APP_NAME, font=FONTS["heading"]).pack(side='left')

        tier = self.license_manager.get_tier_info()
        tier_name = tier['name_ja'] if get_language() == 'ja' else tier['name']

        # 現在のステータス表示
        status_frame = ttk.Frame(frame, style="Card.TFrame")
        status_frame.pack(fill='x', pady=(0, 15))

        ttk.Label(status_frame, text=t('license_current'), font=FONTS["body_bold"]).pack(anchor='w', pady=(10, 5), padx=10)

        status_inner = ttk.Frame(status_frame)
        status_inner.pack(fill='x', padx=10, pady=(0, 10))

        badge_text = f"{tier['badge']}"
        ttk.Label(status_inner, text=badge_text, font=FONTS["body_bold"],
                  foreground=COLOR_PALETTE["brand_primary"]).pack(side='left')

        if self.license_manager.is_activated():
            # 有効期限表示
            expiry_str = self.license_manager.get_expiry_date_str()
            days = self.license_manager.get_days_until_expiry()

            if days is not None:
                if days > 0:
                    status_text = f"  |  {t('license_valid_until', expiry_str)} {t('license_days_remaining', days)}"
                    status_color = COLOR_PALETTE["success"] if days > 30 else COLOR_PALETTE["warning"]
                else:
                    status_text = f"  |  {t('license_status_expired')}"
                    status_color = COLOR_PALETTE["error"]
            else:
                status_text = f"  |  {t('license_perpetual')}"
                status_color = COLOR_PALETTE["success"]

            ttk.Label(status_inner, text=status_text, font=FONTS["small"],
                      foreground=status_color).pack(side='left')

        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=10)

        # フォーマット説明
        ttk.Label(frame, text="形式: INS-SLIDE-{TIER}-XXXX-XXXX-CC", font=FONTS["small"],
                  foreground=COLOR_PALETTE["text_muted"]).pack(anchor='w', pady=(0, 10))

        # メールアドレス入力
        ttk.Label(frame, text=t('license_email'), font=FONTS["body"]).pack(anchor='w')
        email_var = tk.StringVar(value=self.license_manager.license_info.get('email', ''))
        email_entry = ttk.Entry(frame, textvariable=email_var, width=45, font=FONTS["body"])
        email_entry.pack(fill='x', pady=(5, 10))

        # ライセンスキー入力
        ttk.Label(frame, text=t('license_key'), font=FONTS["body"]).pack(anchor='w')
        key_var = tk.StringVar(value=self.license_manager.license_info.get('key', ''))
        key_entry = ttk.Entry(frame, textvariable=key_var, width=45, font=FONTS["body"])
        key_entry.pack(fill='x', pady=(5, 10))

        # エラーメッセージ表示用
        error_var = tk.StringVar()
        error_label = ttk.Label(frame, textvariable=error_var, font=FONTS["small"],
                                foreground=COLOR_PALETTE["error"])
        error_label.pack(anchor='w', pady=(0, 10))

        def activate():
            email = email_var.get().strip()
            key = key_var.get().strip()

            if not email:
                error_var.set(t('license_email_required'))
                return
            if not key:
                error_var.set(t('license_enter_prompt'))
                return

            ok, msg = self.license_manager.activate(email, key)
            if ok:
                messagebox.showinfo(t('dialog_complete'), msg)
                dialog.destroy()
                self._create_layout()
            else:
                error_var.set(msg)

        def deactivate():
            self.license_manager.deactivate()
            messagebox.showinfo(t('dialog_complete'), t('license_deactivated'))
            dialog.destroy()
            self._create_layout()

        def skip_free():
            """Free版として続行"""
            dialog.destroy()

        # ボタンフレーム
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill='x', pady=(10, 0))

        if self.license_manager.is_activated():
            ttk.Button(btn_frame, text=t('btn_deactivate'), command=deactivate).pack(side='left')

        ttk.Button(btn_frame, text=t('btn_activate'), command=activate, style="Accent.TButton").pack(side='left', padx=5)

        if not startup_check:
            ttk.Button(btn_frame, text=t('btn_close'), command=dialog.destroy).pack(side='right')
        else:
            # 起動時はFree版として続行可能
            ttk.Button(btn_frame, text=t('license_continue_free'), command=skip_free).pack(side='right')

        # リンクフレーム
        link_frame = ttk.Frame(frame)
        link_frame.pack(fill='x', pady=(20, 0))

        def open_trial():
            webbrowser.open(SUPPORT_LINKS.get('contact', ''))

        def open_purchase():
            webbrowser.open(SUPPORT_LINKS.get('purchase', ''))

        trial_link = ttk.Label(link_frame, text=t('license_trial_link'), font=FONTS["small"],
                               foreground=COLOR_PALETTE["brand_primary"], cursor="hand2")
        trial_link.pack(side='left')
        trial_link.bind("<Button-1>", lambda e: open_trial())

        ttk.Label(link_frame, text="  |  ", font=FONTS["small"],
                  foreground=COLOR_PALETTE["text_muted"]).pack(side='left')

        purchase_link = ttk.Label(link_frame, text=t('btn_purchase'), font=FONTS["small"],
                                  foreground=COLOR_PALETTE["brand_primary"], cursor="hand2")
        purchase_link.pack(side='left')
        purchase_link.bind("<Button-1>", lambda e: open_purchase())

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
            if not messagebox.askokcancel(t('dialog_confirm_title'), t('dialog_processing_exit')):
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
