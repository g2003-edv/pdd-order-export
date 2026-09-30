@echo off
rem ============================================================
rem  tunnel_to_server.cmd
rem  Keeps a reverse SSH tunnel alive so the nginx vhost on the
rem  server can reach this PC's web UI:
rem      server 127.0.0.1:18765  ->  this PC 127.0.0.1:8765
rem  Started at logon by scheduled task 'PddExportTunnel'.
rem ============================================================
setlocal
cd /d "%~dp0"

set "SSH=C:\Windows\System32\OpenSSH\ssh.exe"
if not exist "%SSH%" set "SSH=ssh"
set "SERVER=root@YOUR_SERVER_IP"          rem 改成你自己的服务器
set "REMOTE_PORT=18765"
set "LOCAL_PORT=8765"

:loop
echo [%date% %time%] connecting %SERVER%  (server:%REMOTE_PORT% -^> pc:%LOCAL_PORT%)
"%SSH%" -o StrictHostKeyChecking=no -o BatchMode=yes -o ServerAliveInterval=30 -o ServerAliveCountMax=3 -o ExitOnForwardFailure=yes -o ConnectTimeout=15 -N -R %REMOTE_PORT%:127.0.0.1:%LOCAL_PORT% %SERVER%
echo [%date% %time%] tunnel exited (code %errorlevel%), retry in 10s
timeout /t 10 /nobreak >nul
goto loop
