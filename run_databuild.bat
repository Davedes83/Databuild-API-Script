@echo off
rem Run the Databuild puller and keep the window open on errors.
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo Python not found on PATH. Install it from https://www.python.org/downloads/
  echo and tick "Add python.exe to PATH" during setup.
  pause
  exit /b 1
)

python databuild_pull.py %*
if errorlevel 1 (
  echo.
  echo Finished with errors above. Fix them, then run this again.
  pause
) else (
  echo.
  echo Done. Files written to your output folder (see .env / console above).
  timeout /t 3
)