@echo off
cd /d "%~dp0"
call .venv\Scripts\activate.bat
python -m excel_report_automation.main
