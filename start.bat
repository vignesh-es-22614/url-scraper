@echo off
cd /d "%~dp0"

set "PY=.venv\Scripts\python.exe"
if not exist "%PY%" (
	echo Virtual environment not found at .venv\Scripts\python.exe
	echo Create it with: py -m venv .venv
	pause
	exit /b 1
)

echo Installing/checking dependencies...
"%PY%" -m pip install -r requirements.txt --quiet
echo.
echo Starting URL Scraper Web UI...
echo Open http://localhost:5000 in your browser
echo Press Ctrl+C to stop.
echo.
"%PY%" app.py
pause
