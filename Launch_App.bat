@echo off
cd /d "%~dp0"
if exist "ExcelImageExtractorPro.exe" (
    start "" "ExcelImageExtractorPro.exe"
) else if exist "dist\ExcelImageExtractorPro\ExcelImageExtractorPro.exe" (
    start "" "dist\ExcelImageExtractorPro\ExcelImageExtractorPro.exe"
) else (
    start "" pythonw app.py
)
exit
