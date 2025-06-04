@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion
echo ====================================
echo 商品登録入力ツール オールインワンリリース
echo ====================================
echo.

:: 現在のバージョンを検出
for /f "tokens=*" %%i in ('findstr /R "CURRENT_VERSION.*=" src\utils\version_checker.py') do set VERSION_LINE=%%i
echo Detected version line: !VERSION_LINE!
for /f "tokens=2 delims==" %%i in ("!VERSION_LINE!") do set CURRENT_VERSION=%%i
set CURRENT_VERSION=!CURRENT_VERSION:"=!
set CURRENT_VERSION=!CURRENT_VERSION: =!
echo Extracted current version: !CURRENT_VERSION!

echo Current Version: !CURRENT_VERSION!
echo.

:: 新しいバージョンを入力
echo Enter new version number (example: 2.1.3):
set /p NEW_VERSION=

if "!NEW_VERSION!"=="" (
    echo ERROR: Version number not entered
    pause
    exit /b 1
)

echo.
echo Updating version from !CURRENT_VERSION! to !NEW_VERSION!
echo Press Y to continue, N to cancel:
set /p CONTINUE=
if /i "!CONTINUE!" neq "Y" (
    echo Cancelled
    pause
    exit /b 0
)

:: Version info
set APP_NAME=ProductRegisterTool
set DISPLAY_NAME=商品登録入力ツール
set ZIP_NAME=ProductRegisterTool-v!NEW_VERSION!.zip
set TAG_NAME=v!NEW_VERSION!
set REPO_OWNER=SEI1026
set REPO_NAME=Product_app

echo.
echo ====================================
echo Starting Build and Release Process
echo ====================================
echo Start time: %DATE% %TIME%
echo Company: 株式会社大宝家具
echo Version: !NEW_VERSION!
echo.

echo [1/10] Environment check...
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found
    pause
    exit /b 1
)

gh --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: GitHub CLI not found
    pause
    exit /b 1
)

gh auth status >nul 2>&1
if errorlevel 1 (
    echo ERROR: GitHub authentication required - run: gh auth login
    pause
    exit /b 1
)

echo Environment check completed

echo [2/10] Required files check...
if not exist "product_app.py" (
    echo ERROR: product_app.py not found
    pause
    exit /b 1
)

if not exist "ProductRegisterTool.spec" (
    echo ERROR: ProductRegisterTool.spec not found
    pause
    exit /b 1
)

echo Required files check completed

echo [3/10] PyInstaller check...
pip show pyinstaller > nul 2>&1
if errorlevel 1 (
    echo Installing PyInstaller...
    pip install pyinstaller
    if errorlevel 1 (
        echo ERROR: PyInstaller installation failed
        pause
        exit /b 1
    )
)

echo [4/10] Cleaning old files...
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build
if exist "!ZIP_NAME!" del /q "!ZIP_NAME!"

echo [5/10] Updating version in source code...
echo Current version detected: !CURRENT_VERSION!
echo New version to set: !NEW_VERSION!

:: PowerShellスクリプトを使用してバージョン更新
powershell -NoProfile -ExecutionPolicy Bypass -File "update_version.ps1" -NewVersion "!NEW_VERSION!"

if errorlevel 1 (
    echo WARNING: Failed to update version in source code
    echo Please manually update src\utils\version_checker.py line 22
    pause
)

:: 更新後のバージョンを確認
echo Verifying version update...
findstr /C:"CURRENT_VERSION" src\utils\version_checker.py

echo [6/10] Building EXE with all data files...
echo This may take several minutes...

:: ビルド前に最終的なバージョン確認
echo Final version check before build:
findstr /C:"CURRENT_VERSION = \"!NEW_VERSION!\"" src\utils\version_checker.py >nul
if errorlevel 1 (
    echo ERROR: Version verification failed - expected CURRENT_VERSION = "!NEW_VERSION!"
    echo Current content:
    findstr /C:"CURRENT_VERSION" src\utils\version_checker.py
    pause
    exit /b 1
) else (
    echo Version verification OK
)

pyinstaller --clean --noconfirm ProductRegisterTool.spec

if errorlevel 1 (
    echo ERROR: Build failed
    pause
    exit /b 1
)

if not exist "dist\ProductRegisterTool.exe" (
    echo ERROR: EXE file not created
    pause
    exit /b 1
)

for %%I in ("dist\ProductRegisterTool.exe") do echo EXE build completed: %%~zI bytes

echo [7/10] Creating ZIP package...
mkdir dist\package_temp
copy "dist\ProductRegisterTool.exe" "dist\package_temp\!DISPLAY_NAME!.exe" >nul

:: Create clean Excel files without test data  
echo Creating clean Excel files for distribution...

:: Copy user-editable Excel files (using clean template as base)
copy "item_template.xlsm" "dist\package_temp\item_manage.xlsm" >nul
copy "item_template.xlsm" "dist\package_temp\item_template.xlsm" >nul

:: Copy entire C# folder structure
xcopy "C#" "dist\package_temp\C#\" /E /I /Y >nul
if errorlevel 1 (
    echo WARNING: Failed to copy C# folder
)

:: Replace C# item.xlsm with clean template
copy "item_template.xlsm" "dist\package_temp\C#\ec_csv_tool\item.xlsm" >nul

:: Create README
(
echo !DISPLAY_NAME! v!NEW_VERSION!
echo ==============================
echo 株式会社大宝家具
echo.
echo ■ 使用方法:
echo   !DISPLAY_NAME!.exe をダブルクリックして起動してください。
echo.
echo ■ 特徴:
echo   - オールインワンの実行ファイル
echo   - 全データファイル内蔵^(CSVマスター、テンプレート、アイコン、C#ツール^)
echo   - 基本動作に追加ファイルは不要
echo   - 自動アップデート機能搭載
echo.
echo ■ 動作環境:
echo   - Windows 10/11 64bit
echo   - .NET Framework 4.5以降^(C#ツール用^)
echo   - メモリ: 4GB以上推奨
echo.
echo Copyright ^(c^) 2025 株式会社大宝家具. All rights reserved.
) > "dist\package_temp\README.txt"

cd dist\package_temp
powershell -NoProfile -ExecutionPolicy Bypass -Command "Compress-Archive -Path '*' -DestinationPath '..\..\!ZIP_NAME!' -Force"
cd ..\..

if not exist "!ZIP_NAME!" (
    echo ERROR: ZIP creation failed
    pause
    exit /b 1
)

for %%I in ("!ZIP_NAME!") do echo ZIP package created: %%~zI bytes

echo [8/10] Updating version.json...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$json = Get-Content 'version.json' | ConvertFrom-Json; " ^
    "$json.version = '!NEW_VERSION!'; " ^
    "$json.release_date = (Get-Date -Format 'yyyy-MM-dd'); " ^
    "$json.download_url = 'https://github.com/!REPO_OWNER!/!REPO_NAME!/releases/download/!TAG_NAME!/!ZIP_NAME!'; " ^
    "$json | ConvertTo-Json -Depth 10 | Set-Content 'version.json'"

if errorlevel 1 (
    echo WARNING: Failed to update version.json
    pause
)

echo [9/10] Creating GitHub release...
gh release view !TAG_NAME! --repo !REPO_OWNER!/!REPO_NAME! >nul 2>&1
if not errorlevel 1 (
    echo WARNING: Release !TAG_NAME! already exists
    echo Delete existing release and create new one?
    choice /c YN /m "[Y]es / [N]o"
    if errorlevel 2 (
        echo Cancelled
        goto cleanup
    )
    echo Deleting existing release...
    gh release delete !TAG_NAME! --repo !REPO_OWNER!/!REPO_NAME! --yes
)

:: Create release notes
(
echo ## !DISPLAY_NAME! v!NEW_VERSION!
echo.
echo ### 🆕 新機能
echo - リアルタイムデータ検証システム
echo - 入力フィールドの即座エラー表示
echo - 商品コード・価格・必須フィールドの検証
echo - 自動保存インジケーター
echo.
echo ### ⚡ 改善点  
echo - UI/UX大幅改善
echo - エラーメッセージ向上
echo - データ品質チェック強化
echo.
echo ### 🐛 バグ修正
echo - Python 3.13互換性修正
echo - 文字エンコーディング問題解決
echo.
echo ### 💾 インストール方法
echo 1. `!ZIP_NAME!` をダウンロード
echo 2. 適当なフォルダに解凍  
echo 3. `!DISPLAY_NAME!.exe` をダブルクリック
echo.
echo ### 💻 動作環境
echo - Windows 10/11 64bit
echo - .NET Framework 4.5以降
echo - メモリ: 4GB以上推奨
echo.
echo ---
echo 🏢 開発: 株式会社大宝家具
echo 📅 リリース: %DATE%
) > release_notes.md

gh release create !TAG_NAME! ^
    --repo !REPO_OWNER!/!REPO_NAME! ^
    --title "!DISPLAY_NAME! v!NEW_VERSION!" ^
    --notes-file "release_notes.md" ^
    "!ZIP_NAME!"

if errorlevel 1 (
    echo ERROR: GitHub release creation failed
    pause
    exit /b 1
)

echo GitHub release created successfully!

echo [10/10] Committing version updates...
git status >nul 2>&1
if not errorlevel 1 (
    echo Checking Git repository status...
    git status --porcelain version.json src\utils\version_checker.py >nul 2>&1
    
    echo Adding version files to Git...
    git add version.json src\utils\version_checker.py
    
    echo Committing version updates...
    git commit -m "chore: update version to v!NEW_VERSION! with auto-update support"
    
    echo Pushing to GitHub...
    git push origin main
    if errorlevel 1 (
        echo WARNING: Failed to push changes to GitHub
        echo Please manually run: git push origin main
    ) else (
        echo ✅ Version updates pushed to GitHub successfully!
        echo Auto-update URL should now be accessible
    )
) else (
    echo Git not available - please manually commit version updates:
    echo 1. git add version.json src\utils\version_checker.py  
    echo 2. git commit -m "chore: update version to v!NEW_VERSION!"
    echo 3. git push origin main
)

echo.
echo ====================================
echo 🎉 Release Completed Successfully!
echo ====================================
echo.
echo 📦 Release: !TAG_NAME!
echo 🔗 URL: https://github.com/!REPO_OWNER!/!REPO_NAME!/releases/tag/!TAG_NAME!
echo 📁 File: !ZIP_NAME!
for %%I in ("!ZIP_NAME!") do echo 📏 Size: %%~zI bytes
echo.
echo ✅ EXE built with all data files
echo ✅ GitHub Release created
echo ✅ Auto-update configuration updated
echo ✅ Ready for distribution
echo.
echo Users will receive automatic update notifications!
echo.

:: Success sound
powershell -c "[Console]::Beep(800, 200); [Console]::Beep(1000, 200); [Console]::Beep(1200, 200)" >nul 2>&1

:cleanup
:: Clean up temporary files
if exist "dist\package_temp" rmdir /s /q "dist\package_temp"
if exist "release_notes.md" del /q "release_notes.md"

echo End time: %DATE% %TIME%
echo.
pause