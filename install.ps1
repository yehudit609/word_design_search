# WordFormatSearch - installer
# Usage (from the project folder):
#   Set-ExecutionPolicy -Scope Process Bypass
#   .\install.ps1

$ErrorActionPreference = "Stop"
$Root     = $PSScriptRoot
$TaskName = "WordFormatSearch Server"
$Port     = 3000
$CertDays = 365
$CertDir  = Join-Path $env:USERPROFILE ".office-addin-dev-certs"
$Vbs      = Join-Path $Root "start-server.vbs"

function Step($msg) { Write-Host ""; Write-Host "==> $msg" -ForegroundColor Cyan }
function Ok($msg)   { Write-Host "    $msg" -ForegroundColor Green }
function Warn($msg) { Write-Host "    $msg" -ForegroundColor Yellow }

Set-Location $Root

# 1. Node.js
Step "Checking Node.js"
$nodeCmd = Get-Command node -ErrorAction SilentlyContinue
if (-not $nodeCmd) {
    throw "Node.js was not found. Install Node.js LTS from https://nodejs.org and run this script again."
}
$NodeExe = $nodeCmd.Source
$nodeVersion = (& $NodeExe --version).TrimStart("v")
if ([int]($nodeVersion.Split(".")[0]) -lt 18) {
    throw "Node.js $nodeVersion is too old. Please install Node.js 18 LTS or newer."
}
Ok "Node.js $nodeVersion ($NodeExe)"

# 2. Dependencies
Step "Installing dependencies (npm install)"
& npm install --no-fund --no-audit
if ($LASTEXITCODE -ne 0) { throw "npm install failed." }
Ok "Dependencies installed"

# 3. HTTPS dev certificate
Step "Checking local HTTPS certificate"
$crtPath = Join-Path $CertDir "localhost.crt"
$needCert = $true
if (Test-Path $crtPath) {
    try {
        $crt = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $crtPath
        $daysLeft = ($crt.NotAfter - (Get-Date)).TotalDays
        if ($daysLeft -gt 30) {
            $needCert = $false
            Ok ("Existing certificate valid until " + $crt.NotAfter.ToString("yyyy-MM-dd"))
        } else {
            Warn "Certificate expires soon - renewing."
            & npx office-addin-dev-certs uninstall | Out-Host
        }
    } catch { Warn "Could not read existing certificate - reinstalling." }
}
if ($needCert) {
    Warn "Windows may ask you to confirm trusting the 'Developer CA for Microsoft Office Add-ins'. Click Yes."
    & npx office-addin-dev-certs install --days $CertDays
    if ($LASTEXITCODE -ne 0) { throw "Certificate installation failed." }
    Ok "Certificate installed (valid for $CertDays days)"
}

# Stop a previously running instance so the new files/certificate are used
Step "Stopping previous server instance (if any)"
$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existing) { Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue }
Get-CimInstance Win32_Process -Filter "Name = 'node.exe'" |
    Where-Object { $_.CommandLine -like "*$Root*server.js*" } |
    ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue; Ok "Stopped PID $($_.ProcessId)" }
Start-Sleep -Milliseconds 500

# Any other process still holding the port (e.g. an older version / "npm start" from another folder)
$holders = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue |
    Select-Object -ExpandProperty OwningProcess -Unique
foreach ($procId in $holders) {
    $p = Get-CimInstance Win32_Process -Filter "ProcessId = $procId"
    if ($p -and $p.Name -eq "node.exe") {
        Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue
        Ok "Stopped old server on port $Port (PID $procId): $($p.CommandLine)"
    } elseif ($p) {
        throw "Port $Port is used by another program: $($p.Name) (PID $procId). Close it and run install.ps1 again."
    }
}
Start-Sleep -Milliseconds 500

# 4 + 6. Scheduled task that starts at logon (hidden window)
Step "Registering scheduled task '$TaskName'"
$userId    = "$env:USERDOMAIN\$env:USERNAME"
$action    = New-ScheduledTaskAction -Execute "wscript.exe" -Argument "`"$Vbs`" `"$NodeExe`"" -WorkingDirectory $Root
$trigger   = New-ScheduledTaskTrigger -AtLogOn -User $userId
$settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
                -StartWhenAvailable -ExecutionTimeLimit ([TimeSpan]::Zero) `
                -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1) -MultipleInstances IgnoreNew
$principal = New-ScheduledTaskPrincipal -UserId $userId -LogonType Interactive -RunLevel Limited

$useStartup = $false
try {
    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings `
        -Principal $principal -Description "Serves the WordFormatSearch Word add-in on https://localhost:$Port" -Force | Out-Null
    Ok "Scheduled task registered (runs at Windows logon)"
} catch {
    Warn "Could not register a scheduled task ($($_.Exception.Message))."
    Warn "Falling back to a shortcut in the Startup folder."
    $useStartup = $true
}

if ($useStartup) {
    $startup  = [Environment]::GetFolderPath("Startup")
    $lnkPath  = Join-Path $startup "WordFormatSearch Server.lnk"
    $shell    = New-Object -ComObject WScript.Shell
    $lnk      = $shell.CreateShortcut($lnkPath)
    $lnk.TargetPath       = "wscript.exe"
    $lnk.Arguments        = "`"$Vbs`" `"$NodeExe`""
    $lnk.WorkingDirectory = $Root
    $lnk.WindowStyle      = 7
    $lnk.Save()
    Ok "Startup shortcut created: $lnkPath"
}

# 5. Start now
Step "Starting the server"
if ($useStartup) {
    Start-Process -FilePath "wscript.exe" -ArgumentList "`"$Vbs`" `"$NodeExe`"" -WorkingDirectory $Root
} else {
    Start-ScheduledTask -TaskName $TaskName
}

$url = "https://localhost:$Port/src/taskpane.html"
$up = $false
for ($i = 0; $i -lt 15; $i++) {
    Start-Sleep -Seconds 1
    try {
        $r = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 3
        if ($r.StatusCode -eq 200) { $up = $true; break }
    } catch { }
}
if ($up) {
    Ok "Server is running: $url"
} else {
    Warn "The server did not respond yet. Check logs\server.log in the project folder."
}

Write-Host ""
Write-Host "Done." -ForegroundColor Green
Write-Host "Next: close Word completely, reopen it, then Home > Add-ins > Shared Folder > WordFormatSearch."
