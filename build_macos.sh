#!/bin/bash
set -e

cd "$(dirname "$0")"

python -m PyInstaller \
  --noconfirm \
  --clean \
  --onedir \
  --windowed \
  --name schedule_to_excel \
  --collect-all rapidocr_onnxruntime \
  --collect-all onnxruntime \
  schedule_to_excel.py

echo
echo "macOS App build finished."
echo "App bundle path: dist/schedule_to_excel.app"
