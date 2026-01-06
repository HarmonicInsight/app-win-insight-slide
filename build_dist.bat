@echo off
chcp 65001 > nul
echo ====================================
echo   InsightSlides 配布用ビルド
echo   (ライセンス情報なし)
echo ====================================
echo.

REM ライセンスファイルのパス
set LICENSE_DIR=%USERPROFILE%\.insightslides
set LICENSE_FILE=%LICENSE_DIR%\license.key
set LICENSE_BACKUP=%LICENSE_DIR%\license.key.backup

REM 依存関係インストール
echo [1/6] 依存関係を確認中...
pip install pyinstaller python-pptx openpyxl pillow tksheet --quiet

REM ライセンスファイルを一時退避
echo [2/6] ライセンスファイルを一時退避中...
if exist "%LICENSE_FILE%" (
    echo    - ライセンスファイルを発見: %LICENSE_FILE%
    copy "%LICENSE_FILE%" "%LICENSE_BACKUP%" > nul
    del "%LICENSE_FILE%"
    echo    - 一時的に削除しました（ビルド後に復元します）
) else (
    echo    - ライセンスファイルなし（配布用ビルドに最適）
)

REM 古いビルドファイルを削除
echo [3/6] 古いビルドファイルをクリーンアップ中...
if exist "build\InsightSlides" rmdir /s /q "build\InsightSlides"
if exist "dist\InsightSlides" rmdir /s /q "dist\InsightSlides"

REM ビルド実行
echo [4/6] 配布用EXEをビルド中...
pyinstaller InsightSlides.spec --noconfirm --clean

REM ライセンスファイルを復元
echo [5/6] ライセンスファイルを復元中...
if exist "%LICENSE_BACKUP%" (
    copy "%LICENSE_BACKUP%" "%LICENSE_FILE%" > nul
    del "%LICENSE_BACKUP%"
    echo    - ライセンスファイルを復元しました
)

REM 結果確認
echo [6/6] ビルド完了確認...
if exist "dist\InsightSlides\InsightSlides.exe" (
    echo.
    echo ====================================
    echo   配布用ビルド成功!
    echo   出力: dist\InsightSlides\
    echo ====================================
    echo.
    dir "dist\InsightSlides\InsightSlides.exe"
    echo.
    echo 【重要】
    echo このビルドはライセンスなし（Free版）の状態です。
    echo ビルドしたマシンで実行すると既存ライセンスを読み込みます。
    echo 配布先では Free版として動作します。
    echo.
    echo 配布フォルダ: dist\InsightSlides\
) else (
    echo.
    echo ====================================
    echo   ビルド失敗
    echo ====================================
)

echo.
pause
