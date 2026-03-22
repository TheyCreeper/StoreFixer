@echo off
cd /d "%~dp0"
echo Current directory: %cd%
echo Looking for: %cd%StoreFixer.exe

if not exist "StoreFixer.exe" (
    echo ERROR: StoreFixer.exe not found!
    echo Please ensure StoreFixer.exe is in the same directory as this script.
    pause
    exit /b 1
)

call RunAsTI.cmd "%~dp0StoreFixer.exe"
pause
