@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo ============================================================
echo   TikTok Monthly Content Planner (7 Omni x 16 Styles)
echo ============================================================
echo.

powershell -NoProfile -Command "try { $c = New-Object Net.Sockets.TcpClient('127.0.0.1', 8765); $c.Close(); exit 0 } catch { exit 1 }"
if %errorlevel%==0 (
    echo [OK] Planner server already running on http://127.0.0.1:8765/planner.html
) else (
    echo [+] Starting Planner Server on port 8765...
    start "TikTok Monthly Planner Server" cmd /k "chcp 65001 >nul && cd /d ""%~dp0"" && python planner_server.py"
    timeout /t 2 /nobreak >nul
)

echo [+] Opening Planner Console in browser...
start "" "http://127.0.0.1:8765/planner.html"
exit /b 0
