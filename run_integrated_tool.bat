@echo off
echo 統合EC管理ツールを起動します...

REM 必要なライブラリのインストール確認
python -c "import win32gui" 2>NUL
if errorlevel 1 (
    echo pywin32をインストールしています...
    pip install pywin32
)

python -c "from PyQt5.QtWebEngineWidgets import QWebEngineView" 2>NUL
if errorlevel 1 (
    echo PyQtWebEngineをインストールしています...
    pip install PyQtWebEngine
)

REM ツール起動
python integrated_ec_tool.py

pause