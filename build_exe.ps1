$ErrorActionPreference = "Stop"

python -m pip install --upgrade pip
python -m pip install --upgrade pyinstaller

python -m PyInstaller `
  --noconfirm `
  --clean `
  --name "Density2excel" `
  --onefile `
  --windowed `
  "main.py"

Write-Output "Build finished: dist\\Density2excel.exe"
