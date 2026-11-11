# WordFormatSearch - uninstaller
# Removes the background server (scheduled task / startup shortcut).
# Does NOT delete the project folder or the HTTPS dev certificate.

$ErrorActionPreference = "Continue"
$Root     = $PSScriptRoot
$TaskName = "WordFormatSearch Server"

$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($task) {
    Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "Scheduled task '$TaskName' removed." -ForegroundColor Green
} else {
    Write-Host "Scheduled task '$TaskName' not found."
}

$lnk = Join-Path ([Environment]::GetFolderPath("Startup")) "WordFormatSearch Server.lnk"
if (Test-Path $lnk) {
    Remove-Item $lnk -Force
    Write-Host "Startup shortcut removed." -ForegroundColor Green
}

Get-CimInstance Win32_Process -Filter "Name = 'node.exe'" |
    Where-Object { $_.CommandLine -like "*$Root*server.js*" } |
    ForEach-Object {
        Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
        Write-Host "Stopped server process (PID $($_.ProcessId))." -ForegroundColor Green
    }

Write-Host ""
Write-Host "The project files were not deleted."
Write-Host "To also remove the dev certificate run: npx office-addin-dev-certs uninstall"
