@echo off
REM Batch processor for TradingView backtest files
REM Drop this file in a folder with your Excel files and double-click to run

echo ========================================
echo TradingView Backtest Analyzer
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python is not installed or not in PATH
    echo Please install Python from python.org
    pause
    exit /b 1
)

echo Processing all Excel files in current directory...
echo.

python "%~dp0tradingview_excel_analyzer.py" --batch "%~dp0"

echo.
echo ========================================
echo Processing complete!
echo ========================================
pause
