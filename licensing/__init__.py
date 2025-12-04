# -*- coding: utf-8 -*-
"""
InsightSlides Licensing Module
ライセンス管理の共通モジュール
"""

from .core import (
    LICENSE_SECRET,
    LICENSE_TIERS,
    KEY_PREFIXES,
    generate_key,
    validate_key,
    get_tier_from_key,
    get_expiration_from_key,
    generate_checksum,
)

__all__ = [
    'LICENSE_SECRET',
    'LICENSE_TIERS',
    'KEY_PREFIXES',
    'generate_key',
    'validate_key',
    'get_tier_from_key',
    'get_expiration_from_key',
    'generate_checksum',
]
