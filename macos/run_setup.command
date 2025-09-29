#!/bin/bash
cd "$(dirname "$0")/.." || exit 1
if [ ! -f .env ]; then
  echo "No .env found. Copy macos/.env.macos.example to .env and fill values first."; exit 1; fi
python3 setup.py
