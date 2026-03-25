@echo off
echo ===================================
echo Starting Master Autofill Generator
echo ===================================
echo.
echo Installing requirements...
pip install -r requirements.txt
echo.
echo Starting Application...
python app.py
pause
