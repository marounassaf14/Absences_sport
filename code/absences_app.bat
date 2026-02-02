@echo off
echo Building EXE...

pyinstaller --onefile --windowed --name "ENSTA_Absence_Management" app.py

REM Move the exe two folders back (root folder)
if exist dist\ENSTA_Absence_Management.exe (
    move /Y dist\ENSTA_Absence_Management.exe ..
    echo EXE moved successfully to root folder.
) else (
    echo ERROR: EXE not found.
)

pause
