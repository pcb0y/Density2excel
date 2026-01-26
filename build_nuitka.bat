@echo off
echo Starting Nuitka build...

:: Create dist directory if it doesn't exist
if not exist dist mkdir dist

:: Install zstandard if missing (needed for Nuitka onefile/standalone)
py -3.13 -m pip install zstandard

:: Run Nuitka
:: --standalone: Create a standalone folder
:: --onefile: Create a single exe (optional, standalone is faster to build and debug)
:: --windows-console-mode=disable: No command window for GUI app
:: --enable-plugin=tk-inter: Enable Tkinter support
:: --windows-icon-from-ico=icon.ico: Use our generated icon
:: --include-data-file=config.ini=config.ini: Include config file
:: --output-dir=dist: Output to dist folder
:: --assume-yes-for-downloads: Allow downloading dependencies (like GCC)

py -3.13 -m nuitka ^
    --standalone ^
    --windows-console-mode=disable ^
    --enable-plugin=tk-inter ^
    --windows-icon-from-ico=icon.ico ^
    --include-data-file=config.ini=config.ini ^
    --output-dir=dist ^
    --remove-output ^
    --assume-yes-for-downloads ^
    main.py

if %errorlevel% neq 0 (
    echo Build failed!
    pause
    exit /b %errorlevel%
)

echo.
echo Build successful!
echo Output is in dist\main.dist\main.exe
pause
