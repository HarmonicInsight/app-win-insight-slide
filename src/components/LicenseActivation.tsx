import { useState } from 'react';
import { useLicenseStore } from '@/stores/licenseStore';
import { TIER_NAMES, PRODUCT_NAMES } from '@insight/license';

interface LicenseActivationProps {
  onClose?: () => void;
}

export function LicenseActivation({ onClose }: LicenseActivationProps) {
  const [inputKey, setInputKey] = useState('');
  const [error, setError] = useState<string | null>(null);

  const {
    licenseKey,
    licenseInfo,
    isActivated,
    limits,
    setLicense,
    clearLicense,
  } = useLicenseStore();

  const handleActivate = () => {
    setError(null);

    if (!inputKey.trim()) {
      setError('ライセンスキーを入力してください');
      return;
    }

    setLicense(inputKey.trim());

    const store = useLicenseStore.getState();
    if (!store.isActivated) {
      setError(store.licenseInfo?.error || 'ライセンスの認証に失敗しました');
    } else {
      setInputKey('');
      onClose?.();
    }
  };

  const handleDeactivate = () => {
    clearLicense();
    setError(null);
  };

  const formatDate = (date: Date | null) => {
    if (!date) return '無期限';
    return new Date(date).toLocaleDateString('ja-JP', {
      year: 'numeric',
      month: 'long',
      day: 'numeric',
    });
  };

  return (
    <div className="license-activation">
      <h2 className="license-activation__title">ライセンス管理</h2>

      {isActivated && licenseInfo ? (
        <div className="license-activation__status">
          <div className="license-activation__status-badge license-activation__status-badge--active">
            アクティブ
          </div>

          <div className="license-activation__info">
            <div className="license-activation__info-row">
              <span className="license-activation__info-label">製品:</span>
              <span className="license-activation__info-value">
                {licenseInfo.product ? PRODUCT_NAMES[licenseInfo.product] : '-'}
              </span>
            </div>

            <div className="license-activation__info-row">
              <span className="license-activation__info-label">プラン:</span>
              <span className="license-activation__info-value">
                {licenseInfo.tier ? TIER_NAMES[licenseInfo.tier] : '-'}
              </span>
            </div>

            <div className="license-activation__info-row">
              <span className="license-activation__info-label">有効期限:</span>
              <span className="license-activation__info-value">
                {formatDate(licenseInfo.expiresAt)}
              </span>
            </div>

            <div className="license-activation__info-row">
              <span className="license-activation__info-label">ライセンスキー:</span>
              <span className="license-activation__info-value license-activation__info-value--key">
                {licenseKey ? `${licenseKey.slice(0, 20)}...` : '-'}
              </span>
            </div>
          </div>

          <div className="license-activation__limits">
            <h3 className="license-activation__limits-title">機能制限</h3>
            <ul className="license-activation__limits-list">
              <li>最大ファイル数: {limits.maxFiles === Infinity ? '無制限' : limits.maxFiles}</li>
              <li>最大レコード数: {limits.maxRecords === Infinity ? '無制限' : limits.maxRecords.toLocaleString()}</li>
              <li>バッチ処理: {limits.batchProcessing ? '有効' : '無効'}</li>
              <li>エクスポート: {limits.export ? '有効' : '無効'}</li>
              <li>クラウド同期: {limits.cloudSync ? '有効' : '無効'}</li>
              <li>優先サポート: {limits.priority ? '有効' : '無効'}</li>
            </ul>
          </div>

          <button
            type="button"
            className="license-activation__button license-activation__button--secondary"
            onClick={handleDeactivate}
          >
            ライセンスを解除
          </button>
        </div>
      ) : (
        <div className="license-activation__form">
          <p className="license-activation__description">
            ライセンスキーを入力して InsightSlide をアクティベートしてください。
          </p>

          <div className="license-activation__input-group">
            <label htmlFor="license-key" className="license-activation__label">
              ライセンスキー
            </label>
            <input
              id="license-key"
              type="text"
              className="license-activation__input"
              value={inputKey}
              onChange={(e) => setInputKey(e.target.value)}
              placeholder="INS-SLIDE-XXX-XXXX-XXXX-XX"
              autoComplete="off"
            />
          </div>

          {error && (
            <div className="license-activation__error">
              {error}
            </div>
          )}

          <div className="license-activation__actions">
            <button
              type="button"
              className="license-activation__button license-activation__button--primary"
              onClick={handleActivate}
            >
              アクティベート
            </button>

            {onClose && (
              <button
                type="button"
                className="license-activation__button license-activation__button--secondary"
                onClick={onClose}
              >
                キャンセル
              </button>
            )}
          </div>

          <p className="license-activation__trial-note">
            ライセンスをお持ちでない場合は、トライアル版として機能制限付きでご利用いただけます。
          </p>
        </div>
      )}
    </div>
  );
}

export default LicenseActivation;
