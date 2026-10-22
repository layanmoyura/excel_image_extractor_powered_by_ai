@echo off
title Build Standalone Executable
cd /d "%~dp0"
echo ==============================================
echo Building ExcelImageExtractorPro.exe...
echo ==============================================
python -m PyInstaller --noconsole --name "ExcelImageExtractorPro" --collect-all customtkinter app.py
if %ERRORLEVEL% EQU 0 (
    echo.
    echo ==============================================
    echo Build Successful!
    echo Executable is located in: dist\ExcelImageExtractorPro\ExcelImageExtractorPro.exe
    echo ==============================================
) else (
    echo.
    echo Build failed with error code %ERRORLEVEL%.
)
pause
