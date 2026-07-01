# Register-GitHubSyncTask.ps1 - Registers a Windows Scheduled Task to sync GitHub repos automatically
#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$ScriptPath  = $null,
    [string]$TaskName    = 'GitHub-Repo-Sync',
    [string]$TaskPath    = '\',
    [switch]$DisableFullSync,
    [switch]$Silent,
    [switch]$Remove
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrEmpty($ScriptPath)) {
    $ScriptPath = Join-Path $PSScriptRoot 'Sync-GitHubRepos.ps1'
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

# Triggers: repeating every 20 minutes indefinitely
$trigger20Min = New-ScheduledTaskTrigger -Once -At '00:00' -RepetitionInterval (New-TimeSpan -Minutes 20)

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
    Trigger   = $trigger20Min
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
