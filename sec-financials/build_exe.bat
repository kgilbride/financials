@echo off
REM ============================================================
REM  Build "SEC Financials.exe" — run this ONCE on a Windows PC.
REM  Double-click this file, or run it from a terminal.
REM  When it finishes, the app is at:  dist\SEC Financials.exe
REM  (in this same folder). Hand that single .exe to the end
REM  user. No Python needed on their machine.
REM ============================================================

cd /d "%~dp0"

echo Installing build tools and dependencies...
python -m pip install --upgrade pip
python -m pip install pyinstaller edgartools openpyxl pandas

echo.
echo Building the app (this can take a few minutes)...
REM Call PyInstaller via "python -m" so it works even when its
REM Scripts folder is not on PATH.
python -m PyInstaller --onefile --windowed --name "SEC Financials" ^
  --collect-all edgar ^
  --collect-all pyarrow ^
  sec_financials_app.py

echo.
if exist "dist\SEC Financials.exe" (
  echo ============================================================
  echo  SUCCESS. Your app is here:
  echo     %CD%\dist\SEC Financials.exe
  echo  Copy that .exe anywhere and double-click to run it.
  echo ============================================================
) else (
  echo ============================================================
  echo  BUILD FAILED - no .exe was produced.
  echo  Scroll up to read the error message for details.
  echo ============================================================
)
pause
