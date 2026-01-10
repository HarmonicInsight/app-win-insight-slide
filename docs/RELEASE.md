# InsightSlides リリース手順

## 概要

GitHub Actions を使用して Windows EXE を自動ビルド・リリースします。

---

## リリース方法

### 方法1: タグによる自動リリース（推奨）

```bash
# 1. 変更をコミット
git add .
git commit -m "Release準備: 機能追加/バグ修正の説明"
git push origin main

# 2. バージョンタグを作成してプッシュ
git tag v1.0.0
git push origin v1.0.0
```

タグをプッシュすると自動的に:
1. Windows環境でEXEをビルド
2. GitHub Releaseを作成
3. EXEファイルをリリースに添付

### 方法2: 手動実行

1. GitHub リポジトリを開く
2. **Actions** タブをクリック
3. 左側から **"Build Windows EXE"** を選択
4. **"Run workflow"** ボタンをクリック
5. オプション:
   - `create_release` にチェック → リリースも作成
   - チェックなし → ビルドのみ（Artifactsからダウンロード可能）

---

## バージョン番号の付け方

### 正式リリース
```
v1.0.0  → メジャーバージョン（大きな変更）
v1.1.0  → マイナーバージョン（機能追加）
v1.1.1  → パッチバージョン（バグ修正）
```

### プレリリース（自動判定）
```
v1.0.0-alpha   → アルファ版
v1.0.0-beta    → ベータ版
v1.0.0-rc1     → リリース候補
```
※ `-alpha`, `-beta`, `-rc` が含まれるタグは自動でプレリリースとしてマークされます

---

## ビルド成果物

| 項目 | 場所 |
|------|------|
| EXEファイル | GitHub Release に添付 |
| ビルドログ | Actions → 該当ワークフロー → ログ |
| Artifacts | Actions → 該当ワークフロー → Artifacts（30日間保存）|

---

## ローカルビルド（開発用）

Windows環境で直接ビルドする場合:

```cmd
cd C:\dev\app-win-insight-slide
build.bat
```

出力: `dist\InsightSlides.exe`

---

## トラブルシューティング

### ビルドが失敗する場合

1. **Actions タブでログを確認**
   - どのステップで失敗したか特定

2. **よくある原因**
   - サブモジュールの同期問題 → `git submodule update --init --recursive`
   - 依存パッケージの問題 → `requirements.txt` を確認
   - Python バージョン → 3.10 を使用

3. **ローカルで再現テスト**
   ```cmd
   pip install -r requirements.txt pyinstaller
   pyinstaller InsightSlides.spec --noconfirm --clean
   ```

### リリースが作成されない場合

- タグが `v` で始まっているか確認（例: `v1.0.0`）
- タグがリモートにプッシュされているか確認: `git push origin <タグ名>`

---

## ワークフローファイル

設定ファイル: `.github/workflows/build.yml`

変更が必要な場合はこのファイルを編集してください。
