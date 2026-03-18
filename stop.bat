@echo off
REM ── Held Time Report — stop script ───────────────────────────────────────
REM  Kills the Flask server process. Run this if you need to stop it manually.

echo Stopping Held Time Report...
for /f "tokens=5" %%a in ('netstat -aon ^| find ":5000" ^| find "LISTENING"') do (
    taskkill /F /PID %%a >nul 2>&1
    echo Stopped process %%a
)
echo Done.
pause
