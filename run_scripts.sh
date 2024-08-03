#!/bin/bash

# Set the script to exit immediately if any command fails
# set -e

# Function to check if Python is installed
function check_and_install_python() {
    # Check if Python is installed by checking the output of python --version
    if ! command -v python &>/dev/null; then
        echo "Python is not installed. Installing Python..."

        # Use PowerShell to install Python from the Microsoft Store or direct download
        powershell -Command "
        $pythonVersion = '3.10'  # Replace with the required version
        if (!(Get-Command python -ErrorAction SilentlyContinue)) {
            Write-Host 'Python is not installed. Installing...'
            $installer = 'https://www.python.org/ftp/python/$pythonVersion/python-$pythonVersion-amd64.exe'
            $downloadPath = \"$env:TEMP\\python-installer.exe\"
            Invoke-WebRequest -Uri $installer -OutFile $downloadPath
            Start-Process -FilePath $downloadPath -ArgumentList '/quiet InstallAllUsers=1 PrependPath=1' -Wait
        } else {
            Write-Host 'Python is already installed.'
        }"

        echo "Python installation completed."
    else
        echo "Python is already installed."
    fi
}

# Step 1: Check and install Python if necessary
check_and_install_python


# Step 1: Install Python dependencies
echo "Installing required Python packages..."

# Ensure pip is updated
pip install --upgrade pip

# Install necessary packages from a requirements file or directly
pip install -r code/requirements.txt

echo "Dependencies installed successfully."

# Step 2: Run the Python script
echo "Running the Python script..."

python code/main.py

echo "Python script executed successfully."

# Wait for user input before exiting
read -p "Press Enter to exit..."
