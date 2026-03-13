@echo off
title Sort Rename Merge Tool

echo =====================================
echo   Sort / Rename / Merge PDF Tool
echo =====================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
echo ERROR: Python is not installed or not in PATH.
echo Please install Python 3.9+ and try again.
pause
exit /b 1
)

REM Run script
echo Running sort_rename_merge.py...
echo.

python sort_rename_merge.py

IF %ERRORLEVEL% NEQ 0 (
echo.
echo Script finished with ERRORS.
) ELSE (
echo.
echo Script finished successfully.
)

echo.
pause
