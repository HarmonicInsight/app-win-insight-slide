@echo off
chcp 65001 > nul
echo ====================================
echo   InsightSlides
echo   完全ビルド＆配布パッケージ作成
echo ====================================
echo.

REM ライセンスファイルのパス
set LICENSE_DIR=%USERPROFILE%\.insightslides
set LICENSE_FILE=%LICENSE_DIR%\license.key
set LICENSE_BACKUP=%LICENSE_DIR%\license.key.backup
set DIST_NAME=InsightSlides_Free

echo このスクリプトは以下を実行します:
echo 1. ライセンスファイルを一時退避
echo 2. クリーンビルド（Free版）
echo 3. ライセンスファイルを復元
echo 4. 配布用ZIPパッケージ作成
echo.
pause

REM ============================================
REM  STEP 1: ライセンス退避
REM ============================================
echo.
echo [1/7] ライセンスファイルを一時退避中...
if exist "%LICENSE_FILE%" (
    echo    - ライセンスファイルを発見: %LICENSE_FILE%
    copy "%LICENSE_FILE%" "%LICENSE_BACKUP%" > nul
    del "%LICENSE_FILE%"
    echo    - 一時的に削除しました
) else (
    echo    - ライセンスファイルなし（配布用ビルドに最適）
)

REM ============================================
REM  STEP 2: 依存関係確認
REM ============================================
echo [2/7] 依存関係を確認中...
pip install pyinstaller python-pptx openpyxl pillow tksheet --quiet

REM ============================================
REM  STEP 3: クリーンアップ
REM ============================================
echo [3/7] 古いビルドファイルをクリーンアップ中...
if exist "build\InsightSlides" rmdir /s /q "build\InsightSlides"
if exist "dist\InsightSlides" rmdir /s /q "dist\InsightSlides"
if exist "release\%DIST_NAME%" rmdir /s /q "release\%DIST_NAME%"
if exist "release\%DIST_NAME%.zip" del /q "release\%DIST_NAME%.zip"

REM ============================================
REM  STEP 4: ビルド実行
REM ============================================
echo [4/7] Free版としてビルド中...
pyinstaller InsightSlides.spec --noconfirm --clean

REM ============================================
REM  STEP 5: ライセンス復元
REM ============================================
echo [5/7] ライセンスファイルを復元中...
if exist "%LICENSE_BACKUP%" (
    copy "%LICENSE_BACKUP%" "%LICENSE_FILE%" > nul
    del "%LICENSE_BACKUP%"
    echo    - ライセンスファイルを復元しました
)

REM ============================================
REM  STEP 6: ビルド確認
REM ============================================
echo [6/7] ビルド結果を確認中...
if not exist "dist\InsightSlides\InsightSlides.exe" (
    echo.
    echo ====================================
    echo   ビルド失敗
    echo ====================================
    pause
    exit /b 1
)
echo    - ビルド成功

REM ============================================
REM  STEP 7: 配布パッケージ作成
REM ============================================
echo [7/7] 配布パッケージを作成中...

mkdir "release\%DIST_NAME%"
xcopy /E /I /Q "dist\InsightSlides\*" "release\%DIST_NAME%\" > nul

REM README作成
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
) > "release\%DIST_NAME%\README.txt"

REM ZIP作成
powershell -Command "Compress-Archive -Path 'release\%DIST_NAME%\*' -DestinationPath 'release\%DIST_NAME%.zip' -Force"

REM ============================================
REM  完了
REM ============================================
echo.
echo.
if exist "release\%DIST_NAME%.zip" (
    echo ====================================
    echo   完全ビルド＆パッケージ作成完了!
    echo ====================================
    echo.
    echo 配布用ZIPファイル: release\%DIST_NAME%.zip
    echo.
    for %%A in ("release\%DIST_NAME%.zip") do echo サイズ: %%~zA バイト ^(%%~zA:~0,-6%% MB^)
    echo.
    echo 【重要】
    echo - このzipを配布してください
    echo - 配布先ではFree版として動作します
    echo - ビルドしたこのPC上で実行するとPRO版として動作しますが、
    echo   それは既存ライセンスが読み込まれるためです
    echo - 配布先では確実にFree版として起動します
    echo.
    echo 【内容物】
    echo - InsightSlides.exe
    echo - _internal\ フォルダ ^(必要なライブラリ^)
    echo - README.txt
    echo.
    explorer "release"
) else (
    echo ====================================
    echo   パッケージ作成失敗
    echo ====================================
)

echo.
pause
