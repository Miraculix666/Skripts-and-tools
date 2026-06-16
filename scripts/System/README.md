# System Configuration and Automation Scripts

This directory contains system utilities, software inventory tools, and local/remote configuration automation.

## Scripts

- **Copy-NetFiles.ps1**: Reliable file transfer utility across network paths.
- **Get-InstalledSoftware.ps1**: Inspects local registries to export detailed lists of installed software.
- **Get-SoftwareCenterList.ps1**: Queries local Software Center catalog.
- **Location_Permissions.ps1**: Verifies and sets file system permissions for directories.
- **Register-GitHubSyncTask.ps1**: Installs scheduled tasks to sync git repositories automatically.
- **Send-EventEmail.ps1**: Monitors Event Viewer logs and fires email reports.
- **Setup-DnsServer.ps1**: Unified script to install DNS Server roles, configure zones, forwarders, and repair existing zones.
- **Sync-GitHubRepos.ps1**: Automates backup, synchronization, pull, and commit actions across local repositories.
- **Sync-PowerPlan.ps1**: Saves, checks, and reapplies standard power plans to clients.
- **Sync-VSCodeSettings.ps1**: Synchronizes settings, extensions, and snippets for VS Code.
- **Test-LogonAndViva.ps1**: Diagnostic tool for checking client logon parameters.
- **Watch-GitHubRepos.ps1**: Lightweight file watcher to detect and trigger sync tasks on repository updates.
