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
rem WARNING: root@ is only for convenience. In production use a dedicated
rem low-privilege account that can only forward ports, and restrict the
rem server's sshd_config with AllowTcpForwarding / PermitOpen. Otherwise a
rem compromised PC gives the attacker direct root access to the server.
set "SERVER=root@YOUR_SERVER_IP"          rem replace with your server
set "REMOTE_PORT=18765"
set "LOCAL_PORT=8765"
rem Attempt cap and backoff: 5s -> 10s -> 20s -> 40s -> 60s (capped)
set "MAX_ATTEMPTS=10"
set "RETRY_DELAY=5"

rem Placeholder not replaced: exit instead of reconnecting to an invalid host
echo %SERVER% | findstr /C:"YOUR_SERVER_IP" >nul
if not errorlevel 1 (
    echo [ERROR] SERVER is still the placeholder %SERVER%, edit this script first.
    exit /b 1
)

set /a ATTEMPT=0

:loop
set /a ATTEMPT+=1
echo [%date% %time%] connecting %SERVER%  (server:%REMOTE_PORT% -^> pc:%LOCAL_PORT%, attempt %ATTEMPT%/%MAX_ATTEMPTS%)
rem StrictHostKeyChecking=accept-new: accept and store the host key on first
rem connect only. Afterwards the key is pinned in %USERPROFILE%\.ssh\known_hosts
rem and a changed key is rejected.
"%SSH%" -o StrictHostKeyChecking=accept-new -o BatchMode=yes -o ServerAliveInterval=30 -o ServerAliveCountMax=3 -o ExitOnForwardFailure=yes -o ConnectTimeout=15 -N -R %REMOTE_PORT%:127.0.0.1:%LOCAL_PORT% %SERVER%
set "CODE=%errorlevel%"

if %ATTEMPT% GEQ %MAX_ATTEMPTS% (
    echo [%date% %time%] tunnel exited ^(code %CODE%^), reached max attempts %MAX_ATTEMPTS%, giving up.
    exit /b %CODE%
)

echo [%date% %time%] tunnel exited ^(code %CODE%^), retry in %RETRY_DELAY%s
timeout /t %RETRY_DELAY% /nobreak >nul
set /a RETRY_DELAY*=2
if %RETRY_DELAY% GTR 60 set "RETRY_DELAY=60"
goto loop
