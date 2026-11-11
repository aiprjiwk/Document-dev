@echo off
title FAT Protocol Generator & Maintenance Tool
echo ========================================================
echo   Launching FAT Protocol Tool...
echo   Streamlit App with Excel Import, Master Rules, and Word Generation
echo ========================================================
cd /d "%~dp0"

echo Opening browser at http://localhost:8503 ...
start "" "http://localhost:8503"

python -m streamlit run app.py --server.port 8503 --server.address localhost --server.headless false

pause
