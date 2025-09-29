@echo off
cd /d %~dp0\..
if not exist .env (
  echo No .env file found. Copy windows\.env.windows.example to .env and fill values.
  pause
  exit /b 1
)
python setup.py
pause
