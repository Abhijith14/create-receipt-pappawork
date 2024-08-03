@echo off
REM Check if Python is installed
python --version >nul 2>&1
IF ERRORLEVEL 1 (
    echo Python is not installed. Installing Python...
    powershell -Command "Start-Process ms-windows-store://pdp/?productid=9PJPW5LDXLZB"
    echo Please install Python from the Microsoft Store and run this script again.
    pause
    exit /b
) ELSE (
    echo Python is installed.
)

REM Install Python dependencies
echo Installing required Python packages...
pip install --upgrade pip
pip install -r code/requirements.txt
REM Run the Python script
echo Running the Python script...
python code/main.py

REM Pause to see the output before the window closes
pause
