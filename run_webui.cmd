@echo off
rem ============================================================
rem  run_webui.cmd - start the PDD export web UI (token protected)
rem  Used by the Windows service / scheduled task for auto start.
rem ============================================================
setlocal
cd /d "%~dp0"

set "PDD_WEB_DATA_DIR=%~dp0data"
rem 未设置 PDD_WEB_ACCESS_PASSWORD 时 webui.py 拒绝监听 0.0.0.0：
rem   - 生产部署请设置访问密码；或
rem   - token-only 部署显式设置 PDD_ALLOW_INSECURE_LOCAL=1（仅监听 127.0.0.1，隧道仍可接入）
set "PDD_WEB_ACCESS_PASSWORD="
set "PDD_WEB_HOST=0.0.0.0"
set "PDD_OPEN_BROWSER=0"
set "PYTHONDONTWRITEBYTECODE=1"
if not defined PDD_WEB_PORT set "PDD_WEB_PORT=8765"

rem --- load the link token from data\link_token.txt ---
set "TOKEN_FILE=%PDD_WEB_DATA_DIR%\link_token.txt"
set "PDD_WEB_LINK_TOKEN="
if exist "%TOKEN_FILE%" for /f "usebackq delims=" %%t in ("%TOKEN_FILE%") do set "PDD_WEB_LINK_TOKEN=%%t"

rem --- FAIL CLOSED: never start the UI without a token ---
rem webui.py also fails closed when PDD_WEB_LINK_TOKEN is empty (all requests
rem get 503), but starting such a process is pointless; stop before launch.
if not defined PDD_WEB_LINK_TOKEN (
  echo [FATAL] link token is empty or "%TOKEN_FILE%" is missing.
  echo [FATAL] refusing to start: the web UI would have NO access control.
  echo [FATAL] create the token file, then start this script again:
  echo [FATAL]   python -c "import secrets;print(secrets.token_urlsafe(24))" ^> "%TOKEN_FILE%"
  exit /b 1
)

echo [info] link token loaded from "%TOKEN_FILE%"
echo [info] listening on %PDD_WEB_HOST%:%PDD_WEB_PORT%

python "%~dp0webui.py"
exit /b %errorlevel%
