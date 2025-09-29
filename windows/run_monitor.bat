@echo off
cd /d %~dp0\..
if not exist .env (
  echo No .env file found. Copy .env.example to .env first.
  pause
  exit /b 1
)
python draft_monitor.py
pause
