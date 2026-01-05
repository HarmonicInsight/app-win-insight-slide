# -*- coding: utf-8 -*-
"""
License Types - ライセンスティア定義
新形式: PPPP-PLAN-YYMM-HASH-SIG1-SIG2
"""
from enum import Enum
from dataclasses import dataclass
from typing import Dict, Any, Optional


class ProductCode(Enum):
    """製品コード（4文字）"""
    INSS = "INSS"  # InsightSlide Standard
    INSP = "INSP"  # InsightSlide Pro


class LicenseTier(str, Enum):
    """ライセンスティア（プラン）"""
    FREE = "FREE"      # 無料版 (機能制限あり)
    TRIAL = "TRIAL"    # トライアル (14日間)
    STD = "STD"        # Standard (年間)
    PRO = "PRO"        # Professional (年間)


class ErrorCode(Enum):
    """エラーコード"""
    E001 = "E001"  # キー形式不正
    E002 = "E002"  # 署名検証失敗
    E003 = "E003"  # メール不一致
    E004 = "E004"  # 期限切れ
    E005 = "E005"  # 製品不一致


ERROR_MESSAGES: Dict[ErrorCode, str] = {
    ErrorCode.E001: "ライセンスキーの形式が正しくありません",
    ErrorCode.E002: "ライセンスキーが無効です",
    ErrorCode.E003: "メールアドレスがライセンスキーと一致しません",
    ErrorCode.E004: "ライセンスの有効期限が切れています",
    ErrorCode.E005: "このライセンスは InsightSlide 用ではありません",
}


# ティア定義
TIERS: Dict[LicenseTier, Dict[str, Any]] = {
    LicenseTier.FREE: {
        "name": "Free",
        "name_ja": "無料版",
        "duration_months": None,
        "duration_days": None,
    },
    LicenseTier.TRIAL: {
        "name": "Trial",
        "name_ja": "トライアル",
        "duration_months": None,
        "duration_days": 14,
    },
    LicenseTier.STD: {
        "name": "Standard",
        "name_ja": "スタンダード",
        "duration_months": 12,
        "duration_days": None,
    },
    LicenseTier.PRO: {
        "name": "Professional",
        "name_ja": "プロフェッショナル",
        "duration_months": 12,
        "duration_days": None,
    },
}


@dataclass
class FeatureLimits:
    """機能制限"""
    update_slide_limit: Optional[int]  # 更新スライド数制限 (None=無制限)
    batch_extract: bool                 # バッチ抽出
    batch_update: bool                  # バッチ更新
    diff_preview: bool                  # 差分プレビュー
    auto_backup: bool                   # 自動バックアップ
    ai_processing: bool                 # AI処理
    speaker_notes: bool                 # スピーカーノート
    font_analysis: bool                 # フォント診断


# 機能制限定義
FEATURE_LIMITS: Dict[LicenseTier, FeatureLimits] = {
    LicenseTier.FREE: FeatureLimits(
        update_slide_limit=3,
        batch_extract=False,
        batch_update=False,
        diff_preview=False,
        auto_backup=False,
        ai_processing=False,
        speaker_notes=False,
        font_analysis=False,
    ),
    LicenseTier.TRIAL: FeatureLimits(
        update_slide_limit=None,
        batch_extract=True,
        batch_update=True,
        diff_preview=True,
        auto_backup=True,
        ai_processing=True,
        speaker_notes=True,
        font_analysis=True,
    ),
    LicenseTier.STD: FeatureLimits(
        update_slide_limit=None,
        batch_extract=True,
        batch_update=False,
        diff_preview=False,
        auto_backup=False,
        ai_processing=True,
        speaker_notes=False,
        font_analysis=False,
    ),
    LicenseTier.PRO: FeatureLimits(
        update_slide_limit=None,
        batch_extract=True,
        batch_update=True,
        diff_preview=True,
        auto_backup=True,
        ai_processing=True,
        speaker_notes=True,
        font_analysis=True,
    ),
}


@dataclass
class ValidationResult:
    """検証結果"""
    valid: bool
    tier: LicenseTier = LicenseTier.FREE
    product_code: Optional[ProductCode] = None
    expires: Optional[str] = None  # YYYY-MM-DD or None
    error_code: Optional[ErrorCode] = None
    error: Optional[str] = None


# 有効な製品コード（InsightSlide用）
VALID_PRODUCT_CODES = [ProductCode.INSS, ProductCode.INSP]

# トライアル期間（日）
TRIAL_DAYS = 14
