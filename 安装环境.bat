@echo off
cd /d %~dp0
where python >nul 2>nul
if errorlevel 1 (
  echo [ERROR] Python 3.10+ was not found. Install Python and add it to PATH.
  pause
  exit /b 1
)
echo [1/2] Installing Python packages...
python -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
if errorlevel 1 goto fail
echo [2/2] Installing Playwright Chromium...
python -m playwright install chromium
if errorlevel 1 goto fail
echo.
echo Installation completed. Run "??????.bat" next.
pause
exit /b 0
:fail
echo.
echo Installation failed. Check the network and error messages above.
pause
exit /b 1
