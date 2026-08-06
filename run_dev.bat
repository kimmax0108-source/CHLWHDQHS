@echo off
setlocal
cd /d %~dp0
set PYTHONPATH=%CD%\src
python main.py
endlocal
