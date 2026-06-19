@echo off
cd /d "%~dp0"
echo Installing/checking dependencies...
pip install -r requirements.txt --quiet
echo.
echo Starting URL Scraper Web UI...
echo Open http://localhost:5000 in your browser
echo Press Ctrl+C to stop.
echo.
python app.py
pause
