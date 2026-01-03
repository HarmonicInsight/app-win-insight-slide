@echo off
chcp 65001 > nul
echo ====================================
echo   InsightSlides EXE Builder
echo ====================================
echo.

REM 依存関係インストール
echo [1/3] 依存関係を確認中...
pip install pyinstaller python-pptx openpyxl pillow --quiet

REM ビルド実行
echo [2/3] EXEをビルド中...
pyinstaller InsightSlides.spec --noconfirm --clean

REM 結果確認
echo [3/3] ビルド完了確認...
if exist "dist\InsightSlides.exe" (
    echo.
    echo ====================================
    echo   ビルド成功!
    echo   出力: dist\InsightSlides.exe
    echo ====================================
    echo.
    dir dist\InsightSlides.exe
) else (
    echo.
    echo ====================================
    echo   ビルド失敗
    echo ====================================
)

echo.
pause
