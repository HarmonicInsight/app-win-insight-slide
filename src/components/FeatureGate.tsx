import { ReactNode } from 'react';
import { useLicenseStore } from '@/stores/licenseStore';
import { FeatureLimits } from '@insight/license';

interface FeatureGateProps {
  /** チェックする機能 */
  feature: keyof FeatureLimits;
  /** 機能が利用可能な場合に表示するコンテンツ */
  children: ReactNode;
  /** 機能が利用不可の場合に表示するコンテンツ（任意） */
  fallback?: ReactNode;
  /** ライセンスがない場合でも表示するか（デフォルト: false） */
  showWhenInactive?: boolean;
}

/**
 * ライセンスの機能制限に基づいてコンテンツを条件付きで表示するコンポーネント
 *
 * @example
 * // クラウド同期機能がある場合のみ表示
 * <FeatureGate feature="cloudSync">
 *   <CloudSyncButton />
 * </FeatureGate>
 *
 * @example
 * // フォールバックコンテンツ付き
 * <FeatureGate
 *   feature="priority"
 *   fallback={<UpgradePrompt />}
 * >
 *   <PrioritySupport />
 * </FeatureGate>
 */
export function FeatureGate({
  feature,
  children,
  fallback = null,
  showWhenInactive = false,
}: FeatureGateProps) {
  const { checkFeature, isActivated } = useLicenseStore();

  // ライセンスが無効で、showWhenInactive が false の場合
  if (!isActivated && !showWhenInactive) {
    return <>{fallback}</>;
  }

  // 機能チェック
  const hasFeature = checkFeature(feature);

  if (hasFeature) {
    return <>{children}</>;
  }

  return <>{fallback}</>;
}

interface LimitGateProps {
  /** チェックする制限 */
  limit: 'maxFiles' | 'maxRecords';
  /** 現在の使用量 */
  currentUsage: number;
  /** 制限内の場合に表示するコンテンツ */
  children: ReactNode;
  /** 制限を超えた場合に表示するコンテンツ（任意） */
  fallback?: ReactNode;
}

/**
 * 数値制限に基づいてコンテンツを条件付きで表示するコンポーネント
 *
 * @example
 * <LimitGate limit="maxFiles" currentUsage={fileCount}>
 *   <AddFileButton />
 * </LimitGate>
 */
export function LimitGate({
  limit,
  currentUsage,
  children,
  fallback = null,
}: LimitGateProps) {
  const { limits, isActivated } = useLicenseStore();

  if (!isActivated) {
    return <>{fallback}</>;
  }

  const maxValue = limits[limit];

  if (typeof maxValue === 'number' && currentUsage < maxValue) {
    return <>{children}</>;
  }

  return <>{fallback}</>;
}

interface UpgradePromptProps {
  /** プロンプトのタイトル */
  title?: string;
  /** プロンプトの説明文 */
  description?: string;
  /** アップグレードボタンのテキスト */
  buttonText?: string;
  /** アップグレードボタンのクリックハンドラ */
  onUpgrade?: () => void;
}

/**
 * アップグレード促進用のプロンプトコンポーネント
 */
export function UpgradePrompt({
  title = 'この機能はご利用いただけません',
  description = 'この機能を利用するには、上位プランへのアップグレードが必要です。',
  buttonText = 'プランを確認',
  onUpgrade,
}: UpgradePromptProps) {
  return (
    <div className="upgrade-prompt">
      <div className="upgrade-prompt__icon">🔒</div>
      <h3 className="upgrade-prompt__title">{title}</h3>
      <p className="upgrade-prompt__description">{description}</p>
      {onUpgrade && (
        <button
          type="button"
          className="upgrade-prompt__button"
          onClick={onUpgrade}
        >
          {buttonText}
        </button>
      )}
    </div>
  );
}

/**
 * 現在のライセンス制限情報を取得するフック
 */
export function useLicenseLimits() {
  const { limits, isActivated, licenseInfo } = useLicenseStore();

  return {
    limits,
    isActivated,
    tier: licenseInfo?.tier ?? null,
    canUseFeature: (feature: keyof FeatureLimits) => {
      if (!isActivated) return false;
      const value = limits[feature];
      return typeof value === 'boolean' ? value : value > 0;
    },
    isWithinLimit: (limit: 'maxFiles' | 'maxRecords', currentUsage: number) => {
      if (!isActivated) return false;
      const maxValue = limits[limit];
      return typeof maxValue === 'number' && currentUsage < maxValue;
    },
    getRemainingQuota: (limit: 'maxFiles' | 'maxRecords', currentUsage: number) => {
      const maxValue = limits[limit];
      if (maxValue === Infinity) return Infinity;
      if (typeof maxValue === 'number') {
        return Math.max(0, maxValue - currentUsage);
      }
      return 0;
    },
  };
}

export default FeatureGate;
