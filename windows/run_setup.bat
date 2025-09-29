@echo off
cd /d %~dp0\..
if not exist .env (
  echo No .env file found. Copy .env.example to .env and fill in your values.
  pause
  exit /b 1
)
python setup.py
pause
