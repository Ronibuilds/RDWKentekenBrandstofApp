@echo off
echo Building RDW Kenteken Checker...

rem Clean previous builds
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

rem Build exe
pyinstaller rdw_kenteken.spec

rem Build installer (met volledig pad naar Inno Setup)
"C:\Users\*\AppData\Local\Programs\Inno Setup 6\ISCC.exe" "setup.iss"

echo Build complete!
pause
