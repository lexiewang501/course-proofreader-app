@echo off
chcp 65001 >nul
title 課程資料校稿小工具 (Word vs PDF 比對系統)
cd /d "%~dp0"

echo =======================================================================
echo          歡迎使用【課程資料校稿小工具】(Word vs PDF 智慧比對)
echo =======================================================================
echo.
echo 正在為您啟動校稿網頁介面，請稍候...
echo.

set AGY_NODE=%LOCALAPPDATA%\Programs\antigravity\Antigravity.exe
if exist "%AGY_NODE%" (
    echo [OK] 偵測到執行環境，正在啟動本機伺服器...
    set ELECTRON_RUN_AS_NODE=1
    start "" http://localhost:8080
    "%AGY_NODE%" server.js
    goto end
)

where node >nul 2>nul
if %errorlevel% equ 0 (
    echo [OK] 偵測到 Node.js，正在啟動伺服器...
    start "" http://localhost:8080
    node server.js
    goto end
)

where python >nul 2>nul
if %errorlevel% equ 0 (
    echo [OK] 偵測到 Python，正在啟動伺服器...
    start "" http://localhost:8080
    python -m http.server 8080
    goto end
)

echo [OK] 直接以預設瀏覽器開啟網頁介面...
start "" "%~dp0index.html"

:end
pause
