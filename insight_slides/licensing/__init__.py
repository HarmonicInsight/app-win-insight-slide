# -*- coding: utf-8 -*-
"""
InsightSlides Licensing Module
ハイブリッド方式: 統一形式 + レガシー形式対応
"""
from .types import (
    LicenseTier,
    TIERS,
    FEATURE_LIMITS,
    FeatureLimits,
    ValidationResult,
    PRODUCT_CODE,
)
from .core import (
    LicenseManager,
    validate_key,
    generate_key,
    generate_trial_key,
)

__all__ = [
    # Types
    "LicenseTier",
    "TIERS",
    "FEATURE_LIMITS",
    "FeatureLimits",
    "ValidationResult",
    "PRODUCT_CODE",
    # Core
    "LicenseManager",
    "validate_key",
    "generate_key",
    "generate_trial_key",
]
