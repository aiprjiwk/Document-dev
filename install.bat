@echo off
setlocal
echo =========================================
echo       Installing Required Programs     
echo =========================================
echo.

:: ตรวจสอบว่ามี Python หรือยัง
python --version >nul 2>&1
if %ERRORLEVEL% == 0 (
    echo [OK] Python is already installed.
    goto install_packages
)

py --version >nul 2>&1
if %ERRORLEVEL% == 0 (
    echo [OK] Python is already installed.
    goto install_packages
)

echo [INFO] Python is not installed. Downloading Python 3.11...
:: ดาวน์โหลด Python 3.11 ผ่าน curl (เพิ่ม --ssl-no-revoke เพื่อแก้ปัญหาดาวน์โหลดไม่ได้)
curl --ssl-no-revoke -L -o python-installer.exe https://www.python.org/ftp/python/3.11.8/python-3.11.8-amd64.exe

if not exist python-installer.exe (
    echo [ERROR] Failed to download Python. Please check your internet or firewall.
    pause
    exit /b
)

echo [INFO] Installing Python... Please wait, this may take a few minutes.
:: ติดตั้ง Python แบบซ่อนหน้าต่าง (Silent) และ Add to PATH ให้อัตโนมัติ
start /wait python-installer.exe /quiet InstallAllUsers=0 PrependPath=1 Include_test=0

echo [OK] Python installation finished.
:: ลบไฟล์ติดตั้งทิ้งเพื่อประหยัดพื้นที่
del python-installer.exe

:: โหลด Environment Variables ใหม่เพื่อให้ระบบรู้จัก Python ทันที
call RefreshEnv.cmd >nul 2>&1

:install_packages
echo.
echo =========================================
echo       Installing Python Packages       
echo =========================================
echo.

:: พยายามติดตั้ง Requirements
echo [INFO] Installing packages from requirements.txt...
py -m pip install --upgrade pip >nul 2>&1
py -m pip install -r requirements.txt
if %ERRORLEVEL% NEQ 0 (
    python -m pip install --upgrade pip >nul 2>&1
    python -m pip install -r requirements.txt
)

echo.
echo =========================================
echo   Installation Completed Successfully!   
echo =========================================
echo You can now run "start.bat" to open the app.
pause