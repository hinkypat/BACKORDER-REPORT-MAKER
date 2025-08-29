@echo off
REM Windows batch file to create executable
echo Building Daily Backorder Report Generator...
echo.

REM Install PyInstaller if not already installed
pip install pyinstaller

REM Build the executable
pyinstaller --onefile --windowed --name="Daily_Backorder_Report_Generator" daily_backorder_app.py

echo.
echo Build complete! 
echo Your executable is in the 'dist' folder:
echo   dist\Daily_Backorder_Report_Generator.exe
echo.
echo You can copy this .exe file anywhere and run it without Python!
pause