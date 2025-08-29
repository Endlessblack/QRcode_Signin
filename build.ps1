param(
  [switch]$Clean
)

# Simple Windows build helper for PyInstaller
$ErrorActionPreference = 'Stop'

Write-Host "== QRcode_Signin build ==" -ForegroundColor Cyan

if (Test-Path .venv\Scripts\Activate.ps1) {
  Write-Host "Activating venv (.venv)" -ForegroundColor DarkCyan
  . .\.venv\Scripts\Activate.ps1
}

if ($Clean) {
  Write-Host "Cleaning build/ and dist/" -ForegroundColor DarkCyan
  Remove-Item -Recurse -Force build, dist -ErrorAction SilentlyContinue
}

if (-not (Get-Command pyinstaller -ErrorAction SilentlyContinue)) {
  Write-Host "PyInstaller not found in PATH. Try: pip install pyinstaller" -ForegroundColor Yellow
  exit 1
}

pyinstaller --noconfirm QRcode_Signin.spec

Write-Host "Build finished. Output in .\\dist\\QRcode_Signin" -ForegroundColor Green

