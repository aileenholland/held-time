@echo off
REM ── Held Time Report — Flask startup script ──────────────────────────────
REM  This file is launched by start_hidden.vbs at Windows login.
REM  Do not run this directly (it will flash a console window).

set PYTHON="C:\Users\Aileen Holland\AppData\Local\Python\pythoncore-3.14-64\python.exe"
set APP_DIR="C:\Users\Aileen Holland\OneDrive - Clearspace Offices\Aileen - Design Group\held-time-report"

cd /d %APP_DIR%
%PYTHON% api\index.py
