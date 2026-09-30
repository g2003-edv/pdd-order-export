$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Data = Join-Path $Root 'data'
$WebPidFile = Join-Path $Data 'webui.pid'
$TunnelPidFile = Join-Path $Data 'cloudflared.pid'

function Stop-ServicePid([string]$PidFile, [string]$Name, [string]$Needle) {
    if (-not (Test-Path -LiteralPath $PidFile)) { return }
    $raw = [string](Get-Content -LiteralPath $PidFile -Raw -ErrorAction SilentlyContinue).Trim()
    if ($raw -notmatch '^\d+$') { Remove-Item -LiteralPath $PidFile -Force -ErrorAction SilentlyContinue; return }
    $procId = [int]$raw
    $proc = Get-Process -Id $procId -ErrorAction SilentlyContinue
    if ($null -eq $proc) {
        Remove-Item -LiteralPath $PidFile -Force -ErrorAction SilentlyContinue
        return
    }
    if ($proc.ProcessName -notlike "$Name*") {
        Write-Host "Skip PID ${procId}: process is $($proc.ProcessName), not $Name."
        Remove-Item -LiteralPath $PidFile -Force -ErrorAction SilentlyContinue
        return
    }
    if ($Needle) {
        try {
            $cim = Get-CimInstance Win32_Process -Filter "ProcessId=$procId" -ErrorAction Stop
            if ($cim.CommandLine -and $cim.CommandLine -notlike "*$Needle*") {
                Write-Host "Skip PID ${procId}: command line does not match this deployment."
                Remove-Item -LiteralPath $PidFile -Force -ErrorAction SilentlyContinue
                return
            }
        } catch {}
    }
    Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $PidFile -Force -ErrorAction SilentlyContinue
    Write-Host "Stopped $Name PID $procId."
}

Stop-ServicePid $TunnelPidFile 'cloudflared' ''
Stop-ServicePid $WebPidFile 'python' 'webui.py'
Write-Host 'Service stopped.'
