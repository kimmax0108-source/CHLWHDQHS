@echo off
setlocal
cd /d %~dp0
python -m pip install -r requirements-dev.txt
python -m pytest
pyinstaller --noconfirm --clean material_document_standardization.spec
if errorlevel 1 exit /b 1
echo.
echo EXE created: dist\material_document_standardization.exe
endlocal
