<#
.SYNOPSIS
    Syncs VS Code settings and extensions to a local repository directory for version control.

.DESCRIPTION
    This script copies your personal VS Code settings.json, keybindings.json, and snippets
    to a designated folder within this repository allowing you to commit and push them to
    GitHub. It also exports a list of your installed extensions.
    You can run it with the -Restore flag on another client to copy the settings back.

.PARAMETER Restore
    If specified, restores the settings from the repository back to your local VS Code configuration.
#>

param (
    [switch]$Restore
)

$repoSettingsDir = Join-Path $PSScriptRoot "VSCode-Settings"
$vscodeUserDir = "$env:APPDATA\Code\User"
$filesToSync = @("settings.json", "keybindings.json", "snippets")

if (-not (Test-Path $repoSettingsDir)) {
    New-Item -ItemType Directory -Path $repoSettingsDir | Out-Null
}

if ($Restore) {
    Write-Host "Restoring VS Code settings from $repoSettingsDir to $vscodeUserDir..." -ForegroundColor Cyan
    foreach ($file in $filesToSync) {
        $sourcePath = "$repoSettingsDir\$file"
        if (Test-Path $sourcePath) {
            Copy-Item $sourcePath "$vscodeUserDir\$file" -Recurse -Force
            Write-Host "Restored $file" -ForegroundColor Green
        }
    }

    $extensionsFile = "$repoSettingsDir\extensions.txt"
    if (Test-Path $extensionsFile) {
        Write-Host "Installing extensions from $extensionsFile..." -ForegroundColor Cyan
        Get-Content $extensionsFile | ForEach-Object {
            if (-not [string]::IsNullOrWhiteSpace($_)) {
                Start-Process -FilePath "code" -ArgumentList "--install-extension $_ --force" -NoNewWindow -Wait
                Write-Host "Installed extension: $_" -ForegroundColor Green
            }
        }
    }
}
else {
    Write-Host "Backing up VS Code settings from $vscodeUserDir to $repoSettingsDir..." -ForegroundColor Cyan
    foreach ($file in $filesToSync) {
        $sourcePath = "$vscodeUserDir\$file"
        if (Test-Path $sourcePath) {
            Copy-Item $sourcePath "$repoSettingsDir\$file" -Recurse -Force
            Write-Host "Backed up $file" -ForegroundColor Green
        }
    }

    $extensionsFile = "$repoSettingsDir\extensions.txt"
    Write-Host "Exporting extension list..." -ForegroundColor Cyan
    # Use cmd to capture code output correctly in some environments
    cmd.exe /c "code --list-extensions" | Out-File -FilePath $extensionsFile -Encoding utf8
    Write-Host "Exported extensions to $extensionsFile" -ForegroundColor Green

    Write-Host "`nDone! You can now commit the VSCode-Settings folder to GitHub." -ForegroundColor Yellow
}
