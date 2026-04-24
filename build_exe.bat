@echo off
setlocal
cd /d "%~dp0"

python -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --windowed ^
  --name schedule_to_excel ^
  --runtime-tmpdir .\pyi_runtime ^
  --collect-all rapidocr_onnxruntime ^
  --collect-all onnxruntime ^
  schedule_to_excel.py

echo.
echo EXE 已建立在 dist\schedule_to_excel.exe
pause
