#!/bin/bash
cd "$(dirname "$0")/.." || exit 1
if [ ! -f .env ]; then
  echo "No .env file. Create one first (copy .env.example to .env)."; exit 1; fi
python3 macos/draft_monitor.py
