@echo off
title Excel Image Extractor Pro
cd /d "%~dp0"
echo Launching Excel Image Extractor Pro...
python app.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Application exited with error code %ERRORLEVEL%.
    pause
)
