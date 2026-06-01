@echo off
cd /d "%~dp0"
where pythonw >nul 2>nul
if %errorlevel%==0 (
    start "" pythonw ".\scripts\product_import_app.py"
) else (
    python ".\scripts\product_import_app.py"
    pause
)
