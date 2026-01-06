@echo off
chcp 65001 > nul
echo ====================================
echo   配布パッケージ作成
echo   InsightSlides
echo ====================================
echo.

REM 配布用フォルダ名
set DIST_NAME=InsightSlides_Free
set RELEASE_DIR=release\%DIST_NAME%

REM ビルド済みexeの確認
if not exist "dist\InsightSlides\InsightSlides.exe" (
    echo エラー: ビルド済みexeが見つかりません
    echo 先に build_dist.bat を実行してください
    echo.
    pause
    exit /b 1
)

echo [1/4] 配布フォルダを準備中...
if exist "release\%DIST_NAME%" rmdir /s /q "release\%DIST_NAME%"
mkdir "release\%DIST_NAME%"

echo [2/4] exeとファイルをコピー中...
xcopy /E /I /Q "dist\InsightSlides\*" "%RELEASE_DIR%\" > nul

echo [3/4] READMEを作成中...
(
echo =====================================
echo   Insight Slides - Free版
echo =====================================
echo.
echo PowerPointテキストの抽出・編集・反映ツール
echo.
echo 【使い方】
echo 1. InsightSlides.exe を実行
echo 2. PPTXファイルをドラッグ^&ドロップ
echo 3. テキストを編集 ^(Excel/JSON出力も可能^)
echo 4. 保存
echo.
echo 【動作環境】
echo - Windows 10/11 ^(64bit^)
echo - .NET Framework 不要
echo.
echo 【Free版の制限】
echo - 3スライドまで更新可能
echo - Excel/JSON出力: 不可
echo - 比較機能: 不可
echo.
echo 【ライセンスアップグレード】
echo Trial版 ^(14日間全機能^): メニュー「ライセンス」から申請
echo Standard/Pro版: ライセンスキー購入で解除
echo.
echo お問い合わせ: support@harmonicinsight.com
echo.
echo by Harmonic Insight
echo © 2025
) > "%RELEASE_DIR%\README.txt"

echo [4/4] ZIPファイルを作成中...
powershell -Command "Compress-Archive -Path 'release\%DIST_NAME%\*' -DestinationPath 'release\%DIST_NAME%.zip' -Force"

if exist "release\%DIST_NAME%.zip" (
    echo.
    echo ====================================
    echo   配布パッケージ作成完了!
    echo ====================================
    echo.
    echo 出力ファイル: release\%DIST_NAME%.zip
    echo.
    dir "release\%DIST_NAME%.zip"
    echo.
    echo 【配布について】
    echo - このzipファイルを配布してください
    echo - 解凍後、InsightSlides.exe を実行
    echo - 配布先ではFree版として動作します
    echo - ライセンスキーは各ユーザーが個別に入力
    echo.
) else (
    echo.
    echo ====================================
    echo   ZIP作成失敗
    echo ====================================
    echo.
)

pause
