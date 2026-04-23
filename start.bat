@echo off
echo ============================================
echo   Visit Plan Approval System
echo ============================================
echo.

:: Install dependencies if not already installed
pip install -r requirements.txt

echo.
echo Starting Flask server...
echo Open your browser at: http://localhost:5000
echo (Also accessible on your network via your IP:5000)
echo.
echo Press CTRL+C to stop the server.
echo.
python app.py
pause
