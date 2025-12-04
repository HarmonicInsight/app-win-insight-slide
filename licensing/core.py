# -*- coding: utf-8 -*-
"""
InsightSlides License Core
ライセンスキーの生成・検証に関する共通ロジック

本体アプリと管理ツールの両方から使用される
"""

import hashlib
import random
import string
from datetime import datetime
from typing import Tuple, Optional

# ============== ライセンス設定 ==============
LICENSE_SECRET = "InsightManagerPro2025"

# キープレフィックスの定義
KEY_PREFIXES = {
    'professional': 'PRO',
    'standard': 'STD',
    'free': 'FRE',
    'trial': 'TRIAL',
}

# ライセンスティア情報
LICENSE_TIERS = {
    'free': {
        'name': 'Free',
        'badge_text': 'Free',
        'badge_style': 'FreeBadge.TLabel',
        'update_slide_limit': 3,
        'format_quality': 'basic',
        'batch_extract': False,
        'batch_update': False,
        'diff_preview': False,
        'auto_backup': False,
        'speaker_notes': False,
        'font_analysis': False,
        'conditional_extract': False,
    },
    'standard': {
        'name': 'Standard',
        'badge_text': '📘 Standard',
        'badge_style': 'StandardBadge.TLabel',
        'price': 2980,
        'update_slide_limit': None,
        'format_quality': 'advanced',
        'batch_extract': True,
        'batch_update': False,
        'diff_preview': False,
        'auto_backup': False,
        'speaker_notes': False,
        'font_analysis': False,
        'conditional_extract': False,
    },
    'professional': {
        'name': 'Professional',
        'badge_text': '⭐ Pro',
        'badge_style': 'ProBadge.TLabel',
        'price': 5980,
        'update_slide_limit': None,
        'format_quality': 'advanced',
        'batch_extract': True,
        'batch_update': True,
        'diff_preview': True,
        'auto_backup': True,
        'speaker_notes': True,
        'font_analysis': True,
        'conditional_extract': True,
    },
}


def generate_checksum(key_body: str) -> str:
    """キーボディからチェックサムを生成"""
    return hashlib.sha256(f"{key_body}{LICENSE_SECRET}".encode()).hexdigest()[:4].upper()


def generate_key(plan: str, key_type: str, expires: str = None) -> str:
    """
    ライセンスキーを生成

    Args:
        plan: プラン名 ('Free', 'Standard', 'Pro')
        key_type: キータイプ ('permanent', 'annual', 'trial')
        expires: 有効期限 ('YYYY-MM-DD' 形式、trial/annual の場合)

    Returns:
        生成されたライセンスキー

    Key formats:
        - permanent: PRO-XXXX-XXXX-XXXX (永続)
        - annual: STD-XXXX-XXXX-2025 (年次、最後が年)
        - trial: TRIAL-XXXXXX-YYYYMMDD (トライアル)
    """
    chars = string.ascii_uppercase + string.digits

    if key_type == "permanent":
        # 永続キー: PRO-XXXX-XXXX-XXXX
        prefix = plan.upper()[:3]
        parts = [''.join(random.choices(chars, k=4)) for _ in range(3)]
        return f"{prefix}-{'-'.join(parts)}"

    elif key_type == "annual":
        # 年間キー: STD-XXXX-XXXX-2025
        prefix = plan.upper()[:3]
        parts = [''.join(random.choices(chars, k=4)) for _ in range(2)]
        year = expires[:4] if expires else datetime.now().strftime("%Y")
        return f"{prefix}-{'-'.join(parts)}-{year}"

    else:  # trial
        # トライアル: TRIAL-XXXXXX-YYYYMMDD
        code = ''.join(random.choices(chars, k=6))
        date = expires.replace("-", "") if expires else datetime.now().strftime("%Y%m%d")
        return f"TRIAL-{code}-{date}"


def validate_key(key: str) -> Tuple[bool, str]:
    """
    ライセンスキーを検証し、有効期限もチェック

    Args:
        key: 検証するライセンスキー

    Returns:
        (is_valid, license_type) のタプル
        - is_valid: キーが有効かどうか
        - license_type: ライセンスタイプ ('free', 'standard', 'professional')
    """
    if not key:
        return False, 'free'

    key = key.strip().upper()
    parts = key.replace(' ', '').split('-')

    # TRIAL-XXXXXX-YYYYMMDD (trial)
    if parts[0] == 'TRIAL' and len(parts) == 3:
        try:
            exp_str = parts[2]
            exp_date = datetime.strptime(exp_str, "%Y%m%d")
            if datetime.now() > exp_date:
                return False, 'free'  # Expired
            return True, 'professional'  # Trial = Pro features
        except:
            return False, 'free'

    # PRO-XXXX-XXXX-XXXX (permanent) or PRO-XXXX-XXXX-2025 (annual)
    if parts[0] in ('PRO', 'STD', 'FRE') and len(parts) == 4:
        license_type = 'professional' if parts[0] == 'PRO' else ('standard' if parts[0] == 'STD' else 'free')
        # Check if last part is a year (annual license)
        if parts[3].isdigit() and len(parts[3]) == 4:
            exp_year = int(parts[3])
            if datetime.now().year > exp_year:
                return False, 'free'  # Expired
        return True, license_type

    # Old format: TYPE-XXXX-XXXX-XXXX-CHECKSUM (5 parts)
    if len(parts) == 5:
        key_body = '-'.join(parts[:4])
        if parts[4] == generate_checksum(key_body):
            license_type = 'professional' if parts[0] == 'PRO' else ('standard' if parts[0] == 'STD' else 'free')
            return True, license_type

    return False, 'free'


def get_tier_from_key(key: str) -> str:
    """キーからライセンスタイプを取得"""
    _, license_type = validate_key(key)
    return license_type


def get_expiration_from_key(key: str) -> Optional[str]:
    """
    キーから有効期限を取得

    Returns:
        'YYYY-MM-DD' 形式の有効期限、永続キーの場合は None
    """
    if not key:
        return None

    key = key.strip().upper()
    parts = key.replace(' ', '').split('-')

    # TRIAL-XXXXXX-YYYYMMDD
    if parts[0] == 'TRIAL' and len(parts) == 3:
        try:
            exp_str = parts[2]
            return f"{exp_str[:4]}-{exp_str[4:6]}-{exp_str[6:8]}"
        except:
            return None

    # PRO-XXXX-XXXX-2025 (annual)
    if len(parts) == 4 and parts[3].isdigit() and len(parts[3]) == 4:
        return f"{parts[3]}-12-31"

    return None  # Permanent or invalid


def get_tier_info(license_type: str) -> dict:
    """ライセンスタイプからティア情報を取得"""
    return LICENSE_TIERS.get(license_type, LICENSE_TIERS['free'])
