# Watch-GitHubRepos.ps1 - Überwacht C:\GitHub auf neue Ordner und löst Workspace-Updates aus
#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$WatchDir = 'C:\GitHub',
    [string]$SyncScript = 'C:\GitHub\Skripts-and-tools\Sync-GitHubRepos.ps1'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Write-Host "👀 Starte Überwachung von '$WatchDir' auf neue Ordner..." -ForegroundColor Cyan

# Create the FileSystemWatcher
$watcher = New-Object IO.FileSystemWatcher $WatchDir, "*.*"
$watcher.IncludeSubdirectories = $false
$watcher.NotifyFilter = [IO.NotifyFilters]::DirectoryName

# Event Action
$action = {
    $path = $Event.SourceEventArgs.FullPath
    $name = $Event.SourceEventArgs.Name
    $changeType = $Event.SourceEventArgs.ChangeType
    $timeStamp = $Event.TimeGenerated

    # Wir reagieren nur auf neue Ordner (keine ignorieren wie .git, .agent)
    if ($name -match '^\.') { return }

    # Check if it's really a directory that was created
    if (Test-Path $path -PathType Container) {
        Write-Host "[$timeStamp] 📁 Neuer Ordner erkannt: $name ($changeType)" -ForegroundColor Green
        Write-Host "🔄 Trigger Sync-Script (nur Workspace Update)..." -ForegroundColor DarkGray

        # We can just call the Sync script silently.
        # It handles updating all.code-workspace and creating the new .code-workspace
        # Setting -Silent to avoid spamming the console
        try {
            & powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File $SyncScript -Silent
            Write-Host "✅ Workspace-Update abgeschlossen für: $name" -ForegroundColor Green
        }
        catch {
            Write-Host "❌ Fehler beim Ausführen des Sync-Scripts: $_" -ForegroundColor Red
        }
    }
}

# Register events
$createdEvent = Register-ObjectEvent $watcher -EventName 'Created' -Action $action
$renamedEvent = Register-ObjectEvent $watcher -EventName 'Renamed' -Action $action

try {
    # Keep the script running
    while ($true) {
        Wait-Event -Timeout 100
    }
}
finally {
    # Cleanup
    Unregister-Event -SourceIdentifier $createdEvent.Name
    Unregister-Event -SourceIdentifier $renamedEvent.Name
    $watcher.Dispose()
    Write-Host "🛑 Überwachung beendet." -ForegroundColor Yellow
}
