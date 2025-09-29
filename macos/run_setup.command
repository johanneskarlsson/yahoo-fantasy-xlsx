#!/bin/bash
cd "$(dirname "$0")/.." || exit 1
if [ ! -f .env ]; then
  echo "No .env found. Copy .env.example to .env and fill in your values first."; exit 1; fi
python3 setup.py
