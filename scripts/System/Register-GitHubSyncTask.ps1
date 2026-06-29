# Register-GitHubSyncTask.ps1 - Registers a Windows Scheduled Task to sync GitHub repos automatically
#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$ScriptPath  = $null,
    [string]$TaskName    = 'GitHub-Repo-Sync',
    [string]$TaskPath    = '\GitHub\',
    [switch]$DisableFullSync,
    [switch]$Silent,
    [switch]$Remove
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrEmpty($ScriptPath)) {
    $ScriptPath = Join-Path $PSScriptRoot 'Sync-GitHubRepos.ps1'
}

# ── Auto-Elevation to Administrator ──────────────────────────────────────────
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "    [WARN] Keine Administratorrechte erkannt. Starte neu als Administrator..." -ForegroundColor Yellow
    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    $paramList = @()
    foreach ($key in $MyInvocation.BoundParameters.Keys) {
        $value = $MyInvocation.BoundParameters[$key]
        if ($value -is [switch]) {
            if ($value) { $paramList += "-$key" }
        } else {
            $paramList += "-$key `"$value`""
        }
    }
    if ($paramList.Count -gt 0) {
        $arguments += " " + ($paramList -join " ")
    }
    Start-Process powershell.exe -ArgumentList $arguments -Verb RunAs
    exit 0
}

if ($Remove) {
    Write-Host "    [DEL] Entferne Scheduled Task '$TaskName' ..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "    [OK] Task entfernt." -ForegroundColor Green
    exit 0
}

Write-Host "    [TASK] Registriere Scheduled Task '$TaskName' ..." -ForegroundColor Cyan

# Action: Run PowerShell with the sync script
$argsList = "-NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$ScriptPath`" -Silent"
if (-not $DisableFullSync) {
    $argsList += " -FullSync"
}

$action = New-ScheduledTaskAction `
    -Execute 'powershell.exe' `
    -Argument $argsList

# Triggers: logon + repeating every 20 minutes indefinitely
$triggerLogon = New-ScheduledTaskTrigger -AtLogOn
$trigger20Min = New-ScheduledTaskTrigger -Once -At '00:00' -RepetitionInterval (New-TimeSpan -Minutes 20) -RepetitionDuration ([System.TimeSpan]::MaxValue)

# Principal: current user, run only when logged on
$principal = New-ScheduledTaskPrincipal `
    -UserId "$env:USERDOMAIN\$env:USERNAME" `
    -LogonType Interactive `
    -RunLevel Limited

# Settings
$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1) `
    -RunOnlyIfNetworkAvailable `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew

# Register
$params = @{
    TaskName  = $TaskName
    TaskPath  = $TaskPath
    Action    = $action
    Trigger   = @($triggerLogon, $trigger20Min)
    Principal = $principal
    Settings  = $settings
    Force     = $true
    Description = "Synchronisiert alle GitHub-Repos nach C:\GitHub\ und aktualisiert VS Code Workspaces."
}

Register-ScheduledTask @params | Out-Null

Write-Host "    [OK] Scheduled Task '$TaskPath$TaskName' erfolgreich registriert." -ForegroundColor Green
Write-Host "   Trigger: alle 20 Minuten + bei Anmeldung" -ForegroundColor DarkGray
Write-Host "   Skript : $ScriptPath" -ForegroundColor DarkGray
if (-not $Silent) {
    Write-Host ""
    Write-Host "   Jetzt manuell ausfuehren? (empfohlen fuer ersten Sync)" -ForegroundColor Yellow
    $run = Read-Host "   [J/n]"
    if ($run -ne 'n' -and $run -ne 'N') {
        Write-Host "    [RUN] Starte Sync ..." -ForegroundColor Cyan
        $runParams = @{}
        if (-not $DisableFullSync) { $runParams['FullSync'] = $true }
        & $ScriptPath @runParams
    }
}
