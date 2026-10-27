# wemeet-watchdog.ps1
# Watchdog: if the wemeet-auto-accept monitor is not running, start it.
# Invoked every 5 minutes by Scheduled Task "WeMeetAutoAcceptWatchdog".
# ASCII-only content to avoid PowerShell 5.1 encoding issues.

$ErrorActionPreference = "SilentlyContinue"
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$monitorScript = Join-Path $dir "wemeet-auto-accept.ps1"
$logFile = Join-Path $dir "wemeet-auto-accept.log"

$mon = Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | Where-Object {
    $_.CommandLine -like '*wemeet-auto-accept.ps1*' -and
    $_.CommandLine -notlike '*Get-CimInstance*' -and
    $_.CommandLine -notlike '*wemeet-watchdog*'
}

if (-not $mon) {
    Start-Process powershell.exe -ArgumentList '-NoProfile','-WindowStyle','Hidden','-ExecutionPolicy','Bypass','-File',"`"$monitorScript`"" -WindowStyle Hidden
    $line = "[{0}] watchdog: monitor not found, restarted" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    Add-Content -Path $logFile -Value $line -Encoding UTF8
}
