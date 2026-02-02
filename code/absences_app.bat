@echo off
echo Building ENSTA Absence Management...

REM Build in ONEDIR mode (fast startup)
pyinstaller --onedir --windowed --clean --name "ENSTA_Absence_Management" app.py

REM Move the generated folder one level up
if exist dist\ENSTA_Absence_Management (
    move /Y dist\ENSTA_Absence_Management ..
    echo Application folder moved successfully.
) else (
    echo ERROR: Build folder not found.
)

REM Cleanup build artifacts
rmdir /S /Q build
rmdir /S /Q dist
del ENSTA_Absence_Management.spec

pause
