# DocForge - 단일 실행파일 빌드 (PyInstaller)
# 사용: powershell -ExecutionPolicy Bypass -File build.ps1
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

python -m PyInstaller --noconfirm --clean `
  --onefile --windowed `
  --name "DocForge" `
  --collect-all docx `
  --exclude-module numba `
  --exclude-module scipy `
  --exclude-module matplotlib `
  --exclude-module IPython `
  --exclude-module notebook `
  --exclude-module PyQt5 `
  --exclude-module PySide2 `
  --exclude-module pandas `
  app.py

Write-Host "`n빌드 완료 -> dist\DocForge.exe"
