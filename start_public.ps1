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
        # 不带 token 探测：服务已监听时必然返回 404（未授权）；只有连接失败才无响应。
        $null = Invoke-WebRequest -UseBasicParsing -Uri "$BaseUrl/" -TimeoutSec 2 -ErrorAction Stop
        $webReady = $true
        break
    } catch {
        if ($_.Exception.Response) { $webReady = $true; break }
    }
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

# 只落盘不含 token 的基础地址：带 token 的完整 URL 一旦被转发/同步到网盘
# 就等于长期有效的钥匙。首次访问所需的 token 单独保存在 data/link_token.txt。
Set-Content -LiteralPath (Join-Path $Data 'public_url.txt') -Value $PublicBase -Encoding ASCII
$AccessHint = @(
    $PublicBase,
    '',
    '首次访问：在上面的地址后附加一次 ?token=<令牌>（令牌见 data/link_token.txt）',
    '浏览器会换取 HttpOnly Cookie 并跳转到不带 token 的地址，之后直接打开基础地址即可。',
    '请不要把带 token 的链接转发给他人。'
) -join [Environment]::NewLine
Set-Content -LiteralPath (Join-Path $Data 'public_access_url.txt') -Value $AccessHint -Encoding UTF8

Write-Host ''
Write-Host '============================================================' -ForegroundColor Green
Write-Host 'Public access is running.' -ForegroundColor Green
Write-Host ''
Write-Host $PublicBase -ForegroundColor Cyan
Write-Host ''
Write-Host 'First visit: append ?token=<link token> once (token: data\link_token.txt).' -ForegroundColor Yellow
Write-Host 'The browser then stores an HttpOnly cookie and drops the token from the URL.' -ForegroundColor Yellow
Write-Host 'Do NOT forward the tokenized link. Share only the base URL above.' -ForegroundColor Yellow
Write-Host 'Keep this computer powered on and connected to the Internet.' -ForegroundColor Yellow
Write-Host '============================================================' -ForegroundColor Green
Write-Host ''
Write-Host "Web PID: $($Web.Id)"
Write-Host "Tunnel PID: $($Tunnel.Id)"
Write-Host "URL file: $(Join-Path $Data 'public_access_url.txt')"
