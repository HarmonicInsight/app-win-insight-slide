# -*- coding: utf-8 -*-
"""
License Core - ライセンス検証・生成
ハイブリッド方式: 統一形式 + レガシー形式対応

統一形式: INS-SLIDE-{TIER}-XXXX-XXXX-CC
レガシー: PRO-XXXX-XXXX-XXXX, STD-XXXX-XXXX-YYYY, TRIAL-XXXXXX-YYYYMMDD
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
LICENSE_SECRET = "InsightManagerPro2025"
UNIFIED_SECRET = "HarmonicInsight2025"


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
                                'is_legacy': result.is_legacy
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
            'is_legacy': result.is_legacy
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
    ライセンスキーを検証 (ハイブリッド方式)

    対応形式:
    1. 統一形式: INS-SLIDE-{TIER}-XXXX-XXXX-CC
    2. レガシー: PRO-XXXX-XXXX-XXXX, TRIAL-XXXXXX-YYYYMMDD
    """
    if not license_key:
        return ValidationResult(valid=False, error="EMPTY_KEY")

    key = license_key.strip().upper()

    # 統一形式 (INS-SLIDE-...)
    if key.startswith("INS-"):
        return _validate_unified_key(key)

    # レガシー形式 (PRO-/STD-/TRIAL-)
    if key.startswith(("PRO-", "STD-", "TRIAL-", "FRE-")):
        return _validate_legacy_key(key)

    return ValidationResult(valid=False, error="INVALID_FORMAT")


def _validate_unified_key(key: str) -> ValidationResult:
    """
    統一形式を検証: INS-SLIDE-{TIER}-XXXX-XXXX-CC
    """
    parts = key.split("-")

    # 形式チェック: INS-SLIDE-TIER-XXXX-XXXX-CC (6パーツ)
    if len(parts) != 6:
        return ValidationResult(valid=False, error="INVALID_UNIFIED_FORMAT")

    prefix, product, tier_str, part1, part2, checksum = parts

    # プレフィックス確認
    if prefix != "INS":
        return ValidationResult(valid=False, error="INVALID_PREFIX")

    # 製品コード確認
    if product != PRODUCT_CODE:
        return ValidationResult(valid=False, error="INVALID_PRODUCT")

    # ティア確認
    try:
        tier = LicenseTier(tier_str)
    except ValueError:
        return ValidationResult(valid=False, error="INVALID_TIER")

    # チェックサム検証
    key_body = f"{prefix}-{product}-{tier_str}-{part1}-{part2}"
    expected_checksum = _generate_checksum(key_body, UNIFIED_SECRET)
    if checksum != expected_checksum:
        return ValidationResult(valid=False, error="INVALID_CHECKSUM")

    # 有効期限計算
    expires = _calculate_expiry_from_parts(tier, part2)

    # 期限チェック
    if expires:
        try:
            exp_date = datetime.strptime(expires, "%Y-%m-%d")
            if datetime.now() > exp_date:
                return ValidationResult(valid=False, tier=LicenseTier.FREE, error="EXPIRED")
        except ValueError:
            pass

    return ValidationResult(valid=True, tier=tier, expires=expires, is_legacy=False)


def _validate_legacy_key(key: str) -> ValidationResult:
    """
    レガシー形式を検証
    - TRIAL-XXXXXX-YYYYMMDD
    - PRO-XXXX-XXXX-XXXX or PRO-XXXX-XXXX-YYYY (年)
    - STD-XXXX-XXXX-XXXX or STD-XXXX-XXXX-YYYY (年)
    """
    parts = key.split("-")

    # TRIAL形式: TRIAL-XXXXXX-YYYYMMDD
    if parts[0] == "TRIAL" and len(parts) == 3:
        try:
            exp_str = parts[2]
            exp_date = datetime.strptime(exp_str, "%Y%m%d")
            if datetime.now() > exp_date:
                return ValidationResult(valid=False, error="EXPIRED", is_legacy=True)
            expires = exp_date.strftime("%Y-%m-%d")
            return ValidationResult(valid=True, tier=LicenseTier.TRIAL, expires=expires, is_legacy=True)
        except ValueError:
            return ValidationResult(valid=False, error="INVALID_TRIAL_FORMAT", is_legacy=True)

    # PRO/STD/FRE形式: TYPE-XXXX-XXXX-XXXX or TYPE-XXXX-XXXX-YYYY
    if parts[0] in ("PRO", "STD", "FRE") and len(parts) == 4:
        tier_map = {"PRO": LicenseTier.PRO, "STD": LicenseTier.STD, "FRE": LicenseTier.FREE}
        tier = tier_map.get(parts[0], LicenseTier.FREE)

        # 年ライセンスチェック (最後のパーツが4桁の年)
        expires = None
        if parts[3].isdigit() and len(parts[3]) == 4:
            exp_year = int(parts[3])
            if datetime.now().year > exp_year:
                return ValidationResult(valid=False, error="EXPIRED", is_legacy=True)
            expires = f"{exp_year}-12-31"

        return ValidationResult(valid=True, tier=tier, expires=expires, is_legacy=True)

    # 旧形式 (5パーツ、チェックサム付き): TYPE-XXXX-XXXX-XXXX-CHECKSUM
    if len(parts) == 5:
        key_body = "-".join(parts[:4])
        checksum = parts[4]
        expected = hashlib.sha256(f"{key_body}{LICENSE_SECRET}".encode()).hexdigest()[:4].upper()

        if checksum == expected:
            tier_map = {"PRO": LicenseTier.PRO, "STD": LicenseTier.STD, "FRE": LicenseTier.FREE}
            tier = tier_map.get(parts[0], LicenseTier.FREE)
            return ValidationResult(valid=True, tier=tier, expires=None, is_legacy=True)

    return ValidationResult(valid=False, error="INVALID_LEGACY_FORMAT", is_legacy=True)


def generate_key(
    tier: LicenseTier,
    expiry_date: Optional[datetime] = None
) -> str:
    """
    統一形式でライセンスキーを生成: INS-SLIDE-{TIER}-XXXX-XXXX-CC
    """
    # ランダムパーツ生成
    chars = string.ascii_uppercase + string.digits
    part1 = ''.join(random.choices(chars, k=4))

    # 有効期限エンコード or ランダム
    if expiry_date:
        # 日付を簡易エンコード
        part2 = expiry_date.strftime("%m%d")[:4].ljust(4, '0')
    else:
        part2 = ''.join(random.choices(chars, k=4))

    # キーボディ
    key_body = f"INS-{PRODUCT_CODE}-{tier.value}-{part1}-{part2}"

    # チェックサム
    checksum = _generate_checksum(key_body, UNIFIED_SECRET)

    return f"{key_body}-{checksum}"


def generate_trial_key(days: int = 14) -> str:
    """トライアルキーを生成"""
    expiry = datetime.now() + timedelta(days=days)
    return generate_key(LicenseTier.TRIAL, expiry)


def _generate_checksum(key_body: str, secret: str) -> str:
    """チェックサムを生成"""
    return hashlib.sha256(f"{key_body}{secret}".encode()).hexdigest()[:2].upper()


def _calculate_expiry_from_parts(tier: LicenseTier, part2: str) -> Optional[str]:
    """パーツから有効期限を計算"""
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
            # 月末調整
            expiry = datetime(new_year, new_month + 1, 1) - timedelta(days=1)
        return expiry.strftime("%Y-%m-%d")

    # ENT: 永久
    return None
