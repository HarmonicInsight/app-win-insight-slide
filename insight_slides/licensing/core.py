# -*- coding: utf-8 -*-
"""
License Core - ライセンス検証・管理
新形式: PPPP-PLAN-YYMM-HASH-SIG1-SIG2

メール紐付け認証: ライセンスキーは発行時のメールアドレスでのみ有効
"""
import hmac
import hashlib
import base64
import json
import os
import re
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional, Tuple

from .types import (
    LicenseTier, ProductCode, ErrorCode, TIERS, FEATURE_LIMITS, FeatureLimits,
    ValidationResult, VALID_PRODUCT_CODES, ERROR_MESSAGES, TRIAL_DAYS
)


# =============================================================================
# 設定
# =============================================================================

# 設定ディレクトリ
CONFIG_DIR = Path.home() / ".insightslides"
LICENSE_FILE = CONFIG_DIR / "license.dat"

# ライセンスキー正規表現
# 形式: PPPP-PLAN-YYMM-HASH-SIG1-SIG2
LICENSE_KEY_REGEX = re.compile(
    r"^(INSS|INSP)-(TRIAL|STD|PRO)-(\d{4})-([A-Z0-9]{4})-([A-Z0-9]{4})-([A-Z0-9]{4})$"
)

# 署名用シークレットキー（環境変数から取得 - 必須）
# セキュリティ上、デフォルト値は設定しない
_SECRET_KEY_RAW = os.environ.get("INSIGHT_LICENSE_SECRET")
_SECRET_KEY = None
_LICENSE_VERIFICATION_AVAILABLE = False

if _SECRET_KEY_RAW:
    _SECRET_KEY = _SECRET_KEY_RAW.encode() if isinstance(_SECRET_KEY_RAW, str) else _SECRET_KEY_RAW
    _LICENSE_VERIFICATION_AVAILABLE = True


# =============================================================================
# 署名・ハッシュ
# =============================================================================

def _generate_email_hash(email: str) -> str:
    """メールアドレスから4文字のハッシュを生成"""
    h = hashlib.sha256(email.lower().strip().encode()).digest()
    return base64.b32encode(h)[:4].decode().upper()


def _generate_signature(data: str) -> str:
    """署名を生成（8文字）- 秘密鍵が必要"""
    if not _LICENSE_VERIFICATION_AVAILABLE or _SECRET_KEY is None:
        raise RuntimeError("License signing not available")
    sig = hmac.new(_SECRET_KEY, data.encode(), hashlib.sha256).digest()
    encoded = base64.b32encode(sig)[:8].decode().upper()
    return encoded


def _verify_signature(data: str, signature: str) -> bool:
    """署名を検証 - 秘密鍵が必要"""
    if not _LICENSE_VERIFICATION_AVAILABLE or _SECRET_KEY is None:
        return False
    try:
        expected = _generate_signature(data)
        return hmac.compare_digest(expected, signature)
    except Exception:
        return False


# =============================================================================
# ライセンスマネージャー
# =============================================================================

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
                    encoded = f.read()
                content = base64.b64decode(encoded).decode()
                data = json.loads(content)

                # 有効期限チェック
                if data.get('expires'):
                    expires = datetime.strptime(data['expires'], "%Y-%m-%d")
                    expires = expires.replace(hour=23, minute=59, second=59)
                    if datetime.now() > expires:
                        return {'type': 'FREE', 'key': '', 'email': '', 'expires': None}

                return {
                    'type': data.get('plan', 'FREE'),
                    'key': data.get('key', ''),
                    'email': data.get('email', ''),
                    'expires': data.get('expires'),
                    'product_code': data.get('productCode'),
                }
            except Exception:
                pass
        return {'type': 'FREE', 'key': '', 'email': '', 'expires': None}

    def _save_license(self, data: dict):
        """ライセンス情報を保存"""
        content = json.dumps(data, ensure_ascii=False)
        encoded = base64.b64encode(content.encode()).decode()
        with open(LICENSE_FILE, 'w', encoding='utf-8') as f:
            f.write(encoded)

    def activate(self, email: str, key: str) -> Tuple[bool, str]:
        """
        ライセンスをアクティベート

        Args:
            email: メールアドレス（キー発行時と同じもの）
            key: ライセンスキー

        Returns:
            (成功フラグ, メッセージ)
        """
        if not email:
            return False, "メールアドレスを入力してください"
        if not key:
            return False, "ライセンスキーを入力してください"

        result = validate_key(email.strip(), key.strip())
        if not result.valid:
            return False, result.error or "無効なライセンスキーです"

        # ライセンス情報を保存
        tier = result.tier
        save_data = {
            'email': email.strip().lower(),
            'key': key.strip().upper(),
            'productCode': result.product_code.value if result.product_code else None,
            'plan': tier.value,
            'expires': result.expires,
            'verifiedAt': datetime.now().isoformat(),
        }
        self._save_license(save_data)

        self.license_info = {
            'type': tier.value,
            'key': key.strip().upper(),
            'email': email.strip().lower(),
            'expires': result.expires,
            'product_code': result.product_code.value if result.product_code else None,
        }

        tier_info = TIERS[tier]
        return True, f"{tier_info['name_ja']}版がアクティベートされました"

    def deactivate(self):
        """ライセンスを解除"""
        self.license_info = {'type': 'FREE', 'key': '', 'email': '', 'expires': None}
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

    def get_email(self) -> str:
        """登録メールアドレスを取得"""
        return self.license_info.get('email', '')


# =============================================================================
# ライセンス検証
# =============================================================================

def validate_key(email: str, license_key: str) -> ValidationResult:
    """
    ライセンスキーを検証

    形式: PPPP-PLAN-YYMM-HASH-SIG1-SIG2
    - PPPP: 製品コード (INSS, INSP)
    - PLAN: プラン (TRIAL, STD, PRO)
    - YYMM: 有効期限（年月）
    - HASH: メールハッシュ（4文字）
    - SIG1-SIG2: HMAC署名（8文字）

    Args:
        email: メールアドレス
        license_key: ライセンスキー

    Returns:
        ValidationResult
    """
    if not email:
        return ValidationResult(valid=False, error="メールアドレスが空です")
    if not license_key:
        return ValidationResult(valid=False, error="キーが空です")

    email = email.strip().lower()
    key = license_key.strip().upper()

    # 1. キー形式チェック
    match = LICENSE_KEY_REGEX.match(key)
    if not match:
        return ValidationResult(
            valid=False,
            error_code=ErrorCode.E001,
            error=ERROR_MESSAGES[ErrorCode.E001]
        )

    product_code_str, plan_str, yymm, email_hash, sig1, sig2 = match.groups()

    try:
        product_code = ProductCode(product_code_str)
        tier = LicenseTier(plan_str)
    except ValueError:
        return ValidationResult(
            valid=False,
            error_code=ErrorCode.E001,
            error=ERROR_MESSAGES[ErrorCode.E001]
        )

    signature = sig1 + sig2

    # 2. 署名検証
    sign_data = f"{product_code_str}-{plan_str}-{yymm}-{email_hash}"
    if not _verify_signature(sign_data, signature):
        return ValidationResult(
            valid=False,
            error_code=ErrorCode.E002,
            error=ERROR_MESSAGES[ErrorCode.E002]
        )

    # 3. メールハッシュ照合
    expected_hash = _generate_email_hash(email)
    if email_hash != expected_hash:
        return ValidationResult(
            valid=False,
            error_code=ErrorCode.E003,
            error=ERROR_MESSAGES[ErrorCode.E003]
        )

    # 4. 有効期限チェック
    try:
        year = 2000 + int(yymm[:2])
        month = int(yymm[2:])
        # 月末日を有効期限とする
        if month == 12:
            expires_dt = datetime(year + 1, 1, 1) - timedelta(days=1)
        else:
            expires_dt = datetime(year, month + 1, 1) - timedelta(days=1)
        expires_dt = expires_dt.replace(hour=23, minute=59, second=59)
        expires = expires_dt.strftime("%Y-%m-%d")
    except ValueError:
        return ValidationResult(
            valid=False,
            error_code=ErrorCode.E001,
            error=ERROR_MESSAGES[ErrorCode.E001]
        )

    if datetime.now() > expires_dt:
        return ValidationResult(
            valid=False,
            tier=tier,
            product_code=product_code,
            expires=expires,
            error_code=ErrorCode.E004,
            error=ERROR_MESSAGES[ErrorCode.E004]
        )

    # 5. 製品コードチェック
    if product_code not in VALID_PRODUCT_CODES:
        return ValidationResult(
            valid=False,
            tier=tier,
            product_code=product_code,
            expires=expires,
            error_code=ErrorCode.E005,
            error=ERROR_MESSAGES[ErrorCode.E005]
        )

    return ValidationResult(
        valid=True,
        tier=tier,
        product_code=product_code,
        expires=expires
    )


# =============================================================================
# ライセンスキー生成（開発者用）
# =============================================================================

def generate_key(product_code: ProductCode, tier: LicenseTier, email: str,
                 expires: Optional[datetime] = None) -> str:
    """
    ライセンスキーを生成

    Args:
        product_code: 製品コード
        tier: プラン
        email: メールアドレス
        expires: 有効期限（省略時は1年後）

    Returns:
        ライセンスキー (PPPP-PLAN-YYMM-HASH-SIG1-SIG2形式)
    """
    if expires is None:
        if tier == LicenseTier.TRIAL:
            expires = datetime.now() + timedelta(days=TRIAL_DAYS)
        else:
            expires = datetime.now() + timedelta(days=365)

    # YYMM形式
    yymm = expires.strftime("%y%m")

    # メールハッシュ
    email_hash = _generate_email_hash(email)

    # 署名データ
    sign_data = f"{product_code.value}-{tier.value}-{yymm}-{email_hash}"

    # 署名生成
    signature = _generate_signature(sign_data)
    sig1, sig2 = signature[:4], signature[4:]

    return f"{product_code.value}-{tier.value}-{yymm}-{email_hash}-{sig1}-{sig2}"


def generate_trial_key(email: str) -> str:
    """トライアルキーを生成（14日間）"""
    return generate_key(ProductCode.INSS, LicenseTier.TRIAL, email)
