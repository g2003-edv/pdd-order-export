$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

$Port = if ($env:PDD_WEB_PORT) { [int]$env:PDD_WEB_PORT } else { 8765 }
$BaseUrl = "http://127.0.0.1:$Port"
$Data = Join-Path $Root 'data'
$Logs = Join-Path $Data 'logs'
New-Item -ItemType Directory -Path $Data -Force | Out-Null
New-Item -ItemType Directory -Path $Logs -Force | Out-Null

function Stop-FromPidFile([string]$PidFile, [string]$ExpectedName, [string]$CommandPart) {
    if (-not (Test-Path -LiteralPath $PidFile)) { return }
    $raw = [string](Get-Content -LiteralPath $PidFile -Raw -ErrorAction SilentlyContinue).Trim()
    if ($raw -notmatch '^\d+$') { Remove-Item -LiteralPath $PidFile -Force -ErrorAction SilentlyContinue; return }
    $procId = [int]$raw
    $proc = Get-Process -Id $procId -ErrorAction SilentlyContinue
    if ($null -eq $proc) { Remove-Item -LiteralPath $PidFile -Force -ErrorAction SilentlyContinue; return }
    if ($proc.ProcessName -notlike "$ExpectedName*") { throw "PID file $PidFile points to unexpected process: $($proc.ProcessName)" }
    try {
        $cim = Get-CimInstance Win32_Process -Filter "ProcessId=$procId" -ErrorAction Stop
        if ($CommandPart -and $cim.CommandLine -and $cim.CommandLine -notlike "*$CommandPart*") {
            throw "PID file $PidFile does not belong to this deployment."
        }
    } catch {
        if ($_.Exception.Message -like 'PID file*') { throw }
    }
    Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $PidFile -Force -ErrorAction SilentlyContinue
    Write-Host "Stopped previous $ExpectedName process: $procId"
}

$WebPidFile = Join-Path $Data 'webui.pid'
$TunnelPidFile = Join-Path $Data 'cloudflared.pid'
Stop-FromPidFile $WebPidFile 'python' 'webui.py'
Stop-FromPidFile $TunnelPidFile 'cloudflared' ''

$portLine = netstat -ano | Select-String -Pattern ":$Port\s+.*LISTENING" | Select-Object -First 1
if ($portLine) {
    throw "Port $Port is already in use. Run 停止服务.bat or close the other program first."
}

$Python = (Get-Command python -ErrorAction Stop).Source
$TokenFile = Join-Path $Data 'link_token.txt'
if (-not (Test-Path -LiteralPath $TokenFile)) {
    $newToken = (& $Python -c "import secrets; print(secrets.token_urlsafe(24))").Trim()
    if (-not $newToken) { throw 'Failed to generate link token.' }
    Set-Content -LiteralPath $TokenFile -Value $newToken -Encoding ASCII
}
$Token = [string](Get-Content -LiteralPath $TokenFile -Raw).Trim()
if (-not $Token) { throw 'Link token is empty.' }

Remove-Item Env:PDD_WEB_ACCESS_PASSWORD -ErrorAction SilentlyContinue
$env:PDD_WEB_LINK_TOKEN = $Token
$env:PDD_WEB_DATA_DIR = $Data
$env:PDD_OPEN_BROWSER = '0'
$env:PDD_WEB_HOST = '0.0.0.0'
$env:PDD_WEB_PORT = [string]$Port
$env:PYTHONDONTWRITEBYTECODE = '1'

$WebOut = Join-Path $Logs 'webui.out.log'
$WebErr = Join-Path $Logs 'webui.err.log'
Remove-Item -LiteralPath $WebOut, $WebErr -Force -ErrorAction SilentlyContinue
$Web = Start-Process -FilePath $Python -ArgumentList @("`"$Root\webui.py`"") -WorkingDirectory $Root -WindowStyle Hidden -RedirectStandardOutput $WebOut -RedirectStandardError $WebErr -PassThru
Set-Content -LiteralPath $WebPidFile -Value $Web.Id -Encoding ASCII

$webReady = $false
for ($i = 0; $i -lt 20; $i++) {
    Start-Sleep -Milliseconds 500
    if ($Web.HasExited) { break }
    try {
        $localUrl = "$BaseUrl/?token=$Token"
        $r = Invoke-WebRequest -UseBasicParsing -Uri $localUrl -TimeoutSec 2
        if ($r.StatusCode -eq 200) { $webReady = $true; break }
    } catch {}
}
if (-not $webReady) {
    if (-not $Web.HasExited) { Stop-Process -Id $Web.Id -Force -ErrorAction SilentlyContinue }
    $details = ''
    foreach ($f in @($WebOut, $WebErr)) {
        if (Test-Path -LiteralPath $f) { $details += [string](Get-Content -LiteralPath $f -Raw -ErrorAction SilentlyContinue) }
    }
    throw "Web server failed to start.`n$details"
}

$Cloudflared = Join-Path $Root 'cloudflared.exe'
if (-not (Test-Path -LiteralPath $Cloudflared)) { throw 'cloudflared.exe is missing.' }
$TunnelOut = Join-Path $Logs 'cloudflared.out.log'
$TunnelErr = Join-Path $Logs 'cloudflared.err.log'
Remove-Item -LiteralPath $TunnelOut, $TunnelErr -Force -ErrorAction SilentlyContinue
$Tunnel = Start-Process -FilePath $Cloudflared -ArgumentList @('tunnel', '--url', $BaseUrl, '--protocol', 'http2', '--no-autoupdate') -WorkingDirectory $Root -WindowStyle Hidden -RedirectStandardOutput $TunnelOut -RedirectStandardError $TunnelErr -PassThru
Set-Content -LiteralPath $TunnelPidFile -Value $Tunnel.Id -Encoding ASCII

$PublicBase = ''
for ($i = 0; $i -lt 60; $i++) {
    Start-Sleep -Seconds 1
    if ($Tunnel.HasExited) { break }
    $text = ''
    foreach ($f in @($TunnelOut, $TunnelErr)) {
        if (Test-Path -LiteralPath $f) { $text += [string](Get-Content -LiteralPath $f -Raw -ErrorAction SilentlyContinue) }
    }
    $m = [regex]::Match($text, 'https://[a-z0-9-]+\.trycloudflare\.com')
    if ($m.Success) { $PublicBase = $m.Value; break }
}
if (-not $PublicBase) {
    if (-not $Tunnel.HasExited) { Stop-Process -Id $Tunnel.Id -Force -ErrorAction SilentlyContinue }
    $details = ''
    foreach ($f in @($TunnelOut, $TunnelErr)) {
        if (Test-Path -LiteralPath $f) { $details += [string](Get-Content -LiteralPath $f -Raw -ErrorAction SilentlyContinue) }
    }
    throw "Cloudflare tunnel failed to start.`n$details"
}

$PublicUrl = "$PublicBase/?token=$Token"
Set-Content -LiteralPath (Join-Path $Data 'public_url.txt') -Value $PublicBase -Encoding ASCII
Set-Content -LiteralPath (Join-Path $Data 'public_access_url.txt') -Value $PublicUrl -Encoding ASCII

Write-Host ''
Write-Host '============================================================' -ForegroundColor Green
Write-Host 'Public access is running.' -ForegroundColor Green
Write-Host ''
Write-Host $PublicUrl -ForegroundColor Cyan
Write-Host ''
Write-Host 'Send this full link, including ?token=...' -ForegroundColor Yellow
Write-Host 'Keep this computer powered on and connected to the Internet.' -ForegroundColor Yellow
Write-Host '============================================================' -ForegroundColor Green
Write-Host ''
Write-Host "Web PID: $($Web.Id)"
Write-Host "Tunnel PID: $($Tunnel.Id)"
Write-Host "URL file: $(Join-Path $Data 'public_access_url.txt')"
