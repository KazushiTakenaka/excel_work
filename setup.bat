@echo off
chcp 65001 > nul
echo ========================================================
echo  Excel & Google Sheets Operation Environment Setup Tool
echo ========================================================
echo.

echo [1/3] Checking Python installation...
python --version > nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed or not in PATH.
    echo Please install Python first according to 01_setup_environment.md.
    pause
    exit /b
)
python --version
echo.

echo [2/3] Creating virtual environment (venv)...
if exist venv (
    echo "venv" folder already exists. Skipping creation.
) else (
    python -m venv venv
    if %errorlevel% neq 0 (
        echo [ERROR] Failed to create venv.
        pause
        exit /b
    )
    echo Created "venv" successfully.
)
echo.

echo [3/3] Installing libraries...
call venv\Scripts\activate.bat
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install libraries.
    pause
    exit /b
)
echo.

echo ========================================================
echo  Setup Completed Successfully!
echo ========================================================
echo.
echo You can now use the virtual environment by running:
echo call venv\Scripts\activate.bat
echo.
pause
