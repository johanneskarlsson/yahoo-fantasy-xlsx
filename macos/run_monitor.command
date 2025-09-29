#!/bin/bash
cd "$(dirname "$0")/.." || exit 1
if [ ! -f .env ]; then
  echo "No .env file. Create one first (copy macos/.env.macos.example)."; exit 1; fi
python3 draft_monitor.py
