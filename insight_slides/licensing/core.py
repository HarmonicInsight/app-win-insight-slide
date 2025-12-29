# -*- coding: utf-8 -*-
"""
License Core - ライセンス検証・生成
統一形式: INS-SLIDE-{TIER}-XXXX-XXXX-CC
"""
import hashlib
import random
import string
import json
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional, Tuple

from .types import (
    LicenseTier, TIERS, FEATURE_LIMITS, FeatureLimits,
    ValidationResult, PRODUCT_CODE
)


# 設定ディレクトリ
CONFIG_DIR = Path.home() / ".insightslides"
LICENSE_FILE = CONFIG_DIR / "license.key"

# シークレット (検証用)
LICENSE_SECRET = "HarmonicInsight2025"


class LicenseManager:
    """ライセンス管理クラス"""

    def __init__(self):
        self._ensure_config_dir()
        self.license_info = self._load_license()

    def _ensure_config_dir(self):
        """設定ディレクトリを作成"""
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)

    def _load_license(self) -> dict:
        """ライセンス情報を読み込み"""
        if LICENSE_FILE.exists():
            try:
                with open(LICENSE_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    if data.get('key'):
                        result = validate_key(data['key'])
                        if result.valid:
                            return {
                                'type': result.tier.value,
                                'key': data['key'],
                                'expires': result.expires,
                            }
            except Exception:
                pass
        return {'type': 'FREE', 'key': '', 'expires': None}

    def _save_license(self, data: dict):
        """ライセンス情報を保存"""
        with open(LICENSE_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def activate(self, key: str) -> Tuple[bool, str]:
        """ライセンスをアクティベート"""
        if not key:
            return False, "ライセンスキーを入力してください"

        result = validate_key(key.strip())
        if not result.valid:
            return False, result.error or "無効なライセンスキーです"

        self.license_info = {
            'type': result.tier.value,
            'key': key.strip().upper(),
            'expires': result.expires,
            'activated': datetime.now().isoformat(),
        }
        self._save_license(self.license_info)

        tier_info = TIERS[result.tier]
        return True, f"{tier_info['name_ja']}版がアクティベートされました"

    def deactivate(self):
        """ライセンスを解除"""
        self.license_info = {'type': 'FREE', 'key': '', 'expires': None}
        if LICENSE_FILE.exists():
            LICENSE_FILE.unlink()

    def get_tier(self) -> LicenseTier:
        """現在のティアを取得"""
        tier_str = self.license_info.get('type', 'FREE')
        try:
            return LicenseTier(tier_str)
        except ValueError:
            return LicenseTier.FREE

    def get_tier_info(self) -> dict:
        """現在のティア情報を取得"""
        return TIERS.get(self.get_tier(), TIERS[LicenseTier.FREE])

    def get_feature_limits(self) -> FeatureLimits:
        """機能制限を取得"""
        return FEATURE_LIMITS.get(self.get_tier(), FEATURE_LIMITS[LicenseTier.FREE])

    def can_use_feature(self, feature: str) -> bool:
        """機能が利用可能かチェック"""
        limits = self.get_feature_limits()
        return getattr(limits, feature, False)

    def get_update_limit(self) -> Optional[int]:
        """更新スライド制限を取得"""
        return self.get_feature_limits().update_slide_limit

    def get_expiration_display(self) -> str:
        """有効期限の表示文字列"""
        expires = self.license_info.get('expires')
        if not expires:
            tier = self.get_tier()
            if tier == LicenseTier.FREE:
                return "-"
            return "永久"

        try:
            exp_date = datetime.strptime(expires, "%Y-%m-%d")
            days_left = (exp_date - datetime.now()).days
            if days_left < 0:
                return "期限切れ"
            return f"残り{days_left}日 ({expires})"
        except ValueError:
            return expires


def validate_key(license_key: str) -> ValidationResult:
    """
    ライセンスキーを検証
    形式: INS-SLIDE-{TIER}-XXXX-XXXX-CC
    """
    if not license_key:
        return ValidationResult(valid=False, error="キーが空です")

    key = license_key.strip().upper()
    parts = key.split("-")

    # 形式チェック: INS-SLIDE-TIER-XXXX-XXXX-CC (6パーツ)
    if len(parts) != 6:
        return ValidationResult(valid=False, error="キー形式が不正です")

    prefix, product, tier_str, part1, part2, checksum = parts

    # プレフィックス確認
    if prefix != "INS":
        return ValidationResult(valid=False, error="プレフィックスが不正です")

    # 製品コード確認
    if product != PRODUCT_CODE:
        return ValidationResult(valid=False, error="製品コードが不正です")

    # ティア確認
    try:
        tier = LicenseTier(tier_str)
    except ValueError:
        return ValidationResult(valid=False, error="ティアが不正です")

    # チェックサム検証
    key_body = f"{prefix}-{product}-{tier_str}-{part1}-{part2}"
    expected_checksum = _generate_checksum(key_body)
    if checksum != expected_checksum:
        return ValidationResult(valid=False, error="チェックサムが不正です")

    # 有効期限計算
    expires = _calculate_expiry(tier)

    # 期限チェック（期限付きの場合）
    if expires:
        try:
            exp_date = datetime.strptime(expires, "%Y-%m-%d")
            if datetime.now() > exp_date:
                return ValidationResult(valid=False, tier=LicenseTier.FREE, error="ライセンスの期限が切れています")
        except ValueError:
            pass

    return ValidationResult(valid=True, tier=tier, expires=expires)


def generate_key(tier: LicenseTier) -> str:
    """
    ライセンスキーを生成
    形式: INS-SLIDE-{TIER}-XXXX-XXXX-CC
    """
    chars = string.ascii_uppercase + string.digits
    part1 = ''.join(random.choices(chars, k=4))
    part2 = ''.join(random.choices(chars, k=4))

    key_body = f"INS-{PRODUCT_CODE}-{tier.value}-{part1}-{part2}"
    checksum = _generate_checksum(key_body)

    return f"{key_body}-{checksum}"


def generate_trial_key() -> str:
    """トライアルキーを生成（14日間）"""
    return generate_key(LicenseTier.TRIAL)


def _generate_checksum(key_body: str) -> str:
    """チェックサムを生成"""
    return hashlib.sha256(f"{key_body}{LICENSE_SECRET}".encode()).hexdigest()[:2].upper()


def _calculate_expiry(tier: LicenseTier) -> Optional[str]:
    """ティアから有効期限を計算"""
    tier_config = TIERS[tier]

    # TRIAL: 14日
    if tier_config.get("duration_days"):
        days = tier_config["duration_days"]
        expiry = datetime.now() + timedelta(days=days)
        return expiry.strftime("%Y-%m-%d")

    # STD/PRO: 12ヶ月
    if tier_config.get("duration_months"):
        months = tier_config["duration_months"]
        now = datetime.now()
        new_month = now.month + months
        new_year = now.year + (new_month - 1) // 12
        new_month = (new_month - 1) % 12 + 1
        try:
            expiry = datetime(new_year, new_month, now.day)
        except ValueError:
            expiry = datetime(new_year, new_month + 1, 1) - timedelta(days=1)
        return expiry.strftime("%Y-%m-%d")

    # FREE/ENT: 永久（期限なし）
    return None
