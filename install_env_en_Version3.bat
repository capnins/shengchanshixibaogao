@echo off
REM Check and install Python dependencies for this project

set PACKAGES=numpy pandas matplotlib pyqt5 openpyxl

echo Checking Python environment...
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Please install Python and ensure it's in PATH!
    pause
    exit /b
)

for %%p in (%PACKAGES%) do (
    python -c "import %%p" >nul 2>&1
    if errorlevel 1 (
        echo [INFO] %%p not found, installing...
        pip install %%p
        if errorlevel 1 (
            echo [ERROR] Failed to install %%p. Please check your network or pip source.
            pause
            exit /b
        )
    ) else (
        echo [OK] %%p is already installed.
    )
)

echo.
echo All dependencies are ready!
pause