@echo off
cd /d "%~dp0"
title SN helper
if not exist ".venv\Scripts\python.exe" (
  echo [ERROR] .venv not found - please set up the Python environment first.
  pause
  exit /b 1
)
set PYTHONIOENCODING=utf-8
".venv\Scripts\python.exe" -u sn_helper.py
pause
