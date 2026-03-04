# Register-GitHubSyncTask.ps1 - Registers a Windows Scheduled Task to sync GitHub repos automatically
#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$ScriptPath  = 'C:\GitHub\Skripts-and-tools\Sync-GitHubRepos.ps1',
    [string]$TaskName    = 'GitHub-Repo-Sync',
    [string]$TaskPath    = '\GitHub\',
    [switch]$Remove
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ($Remove) {
    Write-Host "🗑️  Entferne Scheduled Task '$TaskName' …" -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "✅ Task entfernt." -ForegroundColor Green
    exit 0
}

Write-Host "📅 Registriere Scheduled Task '$TaskName' …" -ForegroundColor Cyan

# Action: Run PowerShell with the sync script
$action = New-ScheduledTaskAction `
    -Execute 'powershell.exe' `
    -Argument "-NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$ScriptPath`" -Silent"

# Triggers: daily at 08:00 + at logon
$triggerDaily  = New-ScheduledTaskTrigger -Daily -At '08:00'
$triggerLogon  = New-ScheduledTaskTrigger -AtLogOn

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
    Trigger   = @($triggerDaily, $triggerLogon)
    Principal = $principal
    Settings  = $settings
    Force     = $true
    Description = "Synchronisiert alle GitHub-Repos nach C:\GitHub\ und aktualisiert VS Code Workspaces."
}

Register-ScheduledTask @params | Out-Null

Write-Host "✅ Scheduled Task '$TaskPath$TaskName' erfolgreich registriert." -ForegroundColor Green
Write-Host "   Trigger: täglich 08:00 + bei Anmeldung" -ForegroundColor DarkGray
Write-Host "   Skript : $ScriptPath" -ForegroundColor DarkGray
Write-Host ""
Write-Host "   Jetzt manuell ausführen? (empfohlen für ersten Sync)" -ForegroundColor Yellow
$run = Read-Host "   [J/n]"
if ($run -ne 'n' -and $run -ne 'N') {
    Write-Host "▶️  Starte Sync …" -ForegroundColor Cyan
    & $ScriptPath
}
