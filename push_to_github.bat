@echo off
title Push to GitHub - layanmoyura
cd /d "%~dp0"

echo ===================================================
echo   Connect and Push to GitHub (layanmoyura)
echo ===================================================
echo.
set REPO_URL=https://github.com/layanmoyura/excel_image_extractor_powered_by_ai.git

echo.
echo Setting remote origin to: %REPO_URL%
git remote remove origin >nul 2>&1
git remote add origin %REPO_URL%

echo.
echo Renaming branch to main...
git branch -M main

echo.
echo Pushing to GitHub...
git push -u origin main

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ===================================================
    echo   Successfully pushed to GitHub!
    echo   URL: https://github.com/layanmoyura/excel_image_extractor_powered_by_ai
    echo ===================================================
) else (
    echo.
    echo Push encountered an error. Please verify your credentials and network connection.
)
pause
