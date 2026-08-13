@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0\.."

if not exist ".venv\Scripts\python.exe" (
  echo [INFO] Creating Python 3.12 virtual environment...
  py -3.12 -m venv .venv || exit /b 1
)

call ".venv\Scripts\activate.bat" || exit /b 1
python -m pip install --upgrade pip || exit /b 1
python -m pip install --editable ".[dev]" || exit /b 1
python -m pytest || exit /b 1
python -m PyInstaller --noconfirm --clean "converter.spec" || exit /b 1

if not exist "release\in" mkdir "release\in"
if not exist "release\out" mkdir "release\out"
if not exist "release\logs" mkdir "release\logs"
copy /y "dist\CSV報表轉換工具.exe" "release\CSV報表轉換工具.exe" >nul || exit /b 1
copy /y "in\CSV檔案放這裡.txt" "release\in\CSV檔案放這裡.txt" >nul || exit /b 1
copy /y "使用說明.txt" "release\使用說明.txt" >nul || exit /b 1

echo [OK] Build completed: release\CSV報表轉換工具.exe
exit /b 0
