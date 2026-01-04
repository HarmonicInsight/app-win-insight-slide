import { useEffect, useState } from 'react';
import { useLicenseStore } from '@/stores/licenseStore';
import { LicenseActivation } from '@/components/LicenseActivation';
import { FeatureGate, UpgradePrompt } from '@/components/FeatureGate';

function App() {
  const [showLicenseModal, setShowLicenseModal] = useState(false);
  const { validateLicense, isActivated, licenseInfo } = useLicenseStore();

  // アプリ起動時にライセンスを検証
  useEffect(() => {
    const result = validateLicense();
    console.log('License validation result:', result);
  }, [validateLicense]);

  return (
    <div className="app">
      <header className="app-header">
        <h1 className="app-header__title">InsightSlide</h1>
        <div className="app-header__actions">
          <button
            type="button"
            className="app-header__license-button"
            onClick={() => setShowLicenseModal(true)}
          >
            {isActivated ? (
              <>
                <span className="app-header__license-status app-header__license-status--active" />
                {licenseInfo?.tier ?? 'アクティブ'}
              </>
            ) : (
              'ライセンス認証'
            )}
          </button>
        </div>
      </header>

      <main className="app-main">
        {/* メインコンテンツはここに配置 */}
        <div className="app-content">
          <p>InsightSlide アプリケーション</p>

          {/* 機能制限の例: クラウド同期 */}
          <FeatureGate
            feature="cloudSync"
            fallback={
              <UpgradePrompt
                title="クラウド同期"
                description="クラウド同期機能を利用するには、PRO プラン以上が必要です。"
                buttonText="アップグレード"
                onUpgrade={() => setShowLicenseModal(true)}
              />
            }
          >
            <div className="cloud-sync-panel">
              <h3>クラウド同期</h3>
              <p>クラウドと同期中...</p>
            </div>
          </FeatureGate>

          {/* 機能制限の例: 優先サポート */}
          <FeatureGate
            feature="priority"
            fallback={
              <UpgradePrompt
                title="優先サポート"
                description="優先サポートを利用するには、PRO プラン以上が必要です。"
              />
            }
          >
            <div className="priority-support">
              <h3>優先サポート</h3>
              <p>優先サポートが有効です。</p>
            </div>
          </FeatureGate>
        </div>
      </main>

      {/* ライセンス認証モーダル */}
      {showLicenseModal && (
        <div className="modal-overlay" onClick={() => setShowLicenseModal(false)}>
          <div className="modal-content" onClick={(e) => e.stopPropagation()}>
            <LicenseActivation onClose={() => setShowLicenseModal(false)} />
          </div>
        </div>
      )}
    </div>
  );
}

export default App;
