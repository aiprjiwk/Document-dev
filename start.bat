@echo off
echo =========================================
echo       Starting PDF OCR Splitter Tool     
echo =========================================
echo.
echo Checking and installing required packages (this might take a moment if it's the first time)...

:: Try using 'py' launcher first, which is standard on Windows
py -m pip install -r requirements.txt >nul 2>&1
if %ERRORLEVEL% == 0 (
    echo Starting the application...
    py -m streamlit run app.py --server.maxUploadSize=4000
    pause
    exit /b
)

:: If 'py' fails, try 'python'
python -m pip install -r requirements.txt >nul 2>&1
if %ERRORLEVEL% == 0 (
    echo Starting the application...
    python -m streamlit run app.py --server.maxUploadSize=4000
    pause
    exit /b
)

:: If 'python' fails, try 'pip' directly
pip install -r requirements.txt >nul 2>&1
if %ERRORLEVEL% == 0 (
    echo Starting the application...
    streamlit run app.py --server.maxUploadSize=4000
    pause
    exit /b
)

echo.
echo ERROR: Python is not recognized!
echo It seems Python was installed without checking the "Add Python to PATH" box.
echo Please double-click python-installer.exe again, choose "Modify", and make sure "Add Python to environment variables" is CHECKED.
pause
