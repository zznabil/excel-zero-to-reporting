@echo off
cd /d "%~dp0"
python -m tools.serve %*
if errorlevel 1 (
  echo.
  echo Course startup failed. Read the error above. Requires Python 3.10 or later.
  echo Read README.md. Do not run as administrator or stop unrelated programs.
  pause
)
