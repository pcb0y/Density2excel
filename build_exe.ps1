$ErrorActionPreference = "Stop"

py -3.8 -m pip install --upgrade pip
py -3.8 -m pip install --upgrade pyinstaller

py -3.8 -m PyInstaller `
  --noconfirm `
  --clean `
  --name "Density2excel" `
  --onedir `
  --windowed `
  --noupx `
  --icon "icon.ico" `
  --add-binary "C:\Windows\System32\vcruntime140.dll;." `
  --add-binary "C:\Windows\System32\vcruntime140_1.dll;." `
  --add-binary "C:\Windows\System32\msvcp140.dll;." `
  --add-binary "C:\Windows\System32\msvcp140_1.dll;." `
  "main.py"

# 复制DLL文件到主目录，确保Windows 7能够找到它们
Copy-Item "dist\Density2excel\_internal\vcruntime140.dll" "dist\Density2excel\" -Force
Copy-Item "dist\Density2excel\_internal\vcruntime140_1.dll" "dist\Density2excel\" -Force
Copy-Item "dist\Density2excel\_internal\msvcp140.dll" "dist\Density2excel\" -Force
Copy-Item "dist\Density2excel\_internal\msvcp140_1.dll" "dist\Density2excel\" -Force

Write-Output "Build finished: dist\Density2excel\Density2excel.exe"
