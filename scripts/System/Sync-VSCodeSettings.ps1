# Sync-VSCodeSettings.ps1 - Sync settings and extensions for standard VS Code and Custom IDE
#Requires -Version 5.1
[CmdletBinding()]
param (
    [switch]$Restore,
    [switch]$Backup
)

$ErrorActionPreference = 'Stop'

# Define path configurations
$repoBaseSettingsDir = Join-Path $PSScriptRoot "VSCode-Settings"

$configs = @(
    @{
        Name           = "VSCode"
        UserDir        = "$env:APPDATA\Code\User"
        RepoDir        = Join-Path $repoBaseSettingsDir "VSCode"
        CliCommand     = "code"
        IsCmd          = $false
    },
    @{
        Name           = "CustomIDE"
        UserDir        = "$env:APPDATA\Antigravity IDE\User"
        RepoDir        = Join-Path $repoBaseSettingsDir "AntigravityIDE"
        CliCommand     = "$env:LOCALAPPDATA\Programs\Antigravity IDE\bin\antigravity-ide.cmd"
        IsCmd          = $true
    }
)

$filesToSync = @("settings.json", "keybindings.json", "snippets")

# Ensure base settings directory exists
if (-not (Test-Path $repoBaseSettingsDir)) {
    New-Item -ItemType Directory -Path $repoBaseSettingsDir | Out-Null
}

function Get-Extensions {
    param(
        [string]$CliCommand,
        [bool]$IsCmd
    )
    
    # We use a unique temporary file to capture stdout
    $guid = [Guid]::NewGuid().Guid
    $tempFile = Join-Path $env:TEMP "exts-$($pid)-$guid.txt"
    
    try {
        # Check if command is available
        $cmdAvailable = $false
        if ($IsCmd) {
            if (Test-Path $CliCommand) { $cmdAvailable = $true }
        } else {
            if (Get-Command $CliCommand -ErrorAction SilentlyContinue) { $cmdAvailable = $true }
        }
        
        if (-not $cmdAvailable) {
            return $null
        }
        
        # Launch process and redirect output
        $proc = Start-Process -FilePath $CliCommand -ArgumentList "--list-extensions" -NoNewWindow -Wait -PassThru -RedirectStandardOutput $tempFile -ErrorAction SilentlyContinue
        
        if (Test-Path $tempFile) {
            $lines = Get-Content $tempFile -ErrorAction SilentlyContinue
            $cleanExts = @()
            if ($null -ne $lines) {
                foreach ($line in $lines) {
                    if ($null -ne $line) {
                        $trimmed = $line.Trim()
                        # Ignore logs starting with [ or containing space or empty
                        if (-not [string]::IsNullOrWhiteSpace($trimmed) -and $trimmed -notmatch '^\[') {
                            $cleanExts += $trimmed
                        }
                    }
                }
            }
            return $cleanExts
        }
    }
    catch {
        Write-Verbose "Failed to run extensions command: $_"
    }
    finally {
        if (Test-Path $tempFile) {
            Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
        }
    }
    return @()
}

# Determine default action (Backup if Restore is not specified)
$actionRestore = $Restore
if (-not $Restore -and -not $Backup) {
    # If neither is specified, default to backup mode
    $actionRestore = $false
}

if ($actionRestore) {
    Write-Host "=== Restoring Editor Settings & Extensions ===" -ForegroundColor Cyan
    
    foreach ($config in $configs) {
        $name = $config.Name
        $userDir = $config.UserDir
        $repoDir = $config.RepoDir
        $cli = $config.CliCommand
        $isCmd = $config.IsCmd
        
        Write-Host ""
        Write-Host "Processing $name..." -ForegroundColor Cyan
        
        if (-not (Test-Path $repoDir)) {
            Write-Host "No backup directory found for $name at $repoDir. Skipping." -ForegroundColor Yellow
            continue
        }
        
        # Ensure user config directory exists
        if (-not (Test-Path $userDir)) {
            New-Item -ItemType Directory -Path $userDir -Force | Out-Null
        }
        
        # 1. Restore files
        foreach ($file in $filesToSync) {
            $sourcePath = Join-Path $repoDir $file
            if (Test-Path $sourcePath) {
                $targetPath = Join-Path $userDir $file
                Copy-Item -Path $sourcePath -Destination $targetPath -Recurse -Force
                Write-Host "  Restored file/folder: $file" -ForegroundColor Green
            }
        }
        
        # 2. Restore/Install extensions
        $extensionsFile = Join-Path $repoDir "extensions.txt"
        if (Test-Path $extensionsFile) {
            $wantedExts = Get-Content $extensionsFile | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() }
            if ($wantedExts.Count -eq 0) {
                Write-Host "  No extensions listed in backup." -ForegroundColor Gray
                continue
            }
            
            # Retrieve currently installed extensions to avoid redundant runs
            Write-Host "  Fetching currently installed extensions..." -ForegroundColor DarkGray
            $installedExts = Get-Extensions -CliCommand $cli -IsCmd $isCmd
            
            if ($null -eq $installedExts) {
                Write-Host "  [WARN] CLI utility for $name ($cli) is not available. Skipping extension installation." -ForegroundColor Yellow
                continue
            }
            
            # Normalize installed extensions to lowercase
            $installedExtsLower = $installedExts | ForEach-Object { $_.ToLower() }
            
            foreach ($ext in $wantedExts) {
                $extLower = $ext.ToLower()
                if ($installedExtsLower -contains $extLower) {
                    Write-Host "  Extension already installed: $ext" -ForegroundColor Gray
                } else {
                    Write-Host "  Installing missing extension: $ext..." -ForegroundColor Yellow
                    try {
                        $installProcess = Start-Process -FilePath $cli -ArgumentList "--install-extension $ext --force" -NoNewWindow -Wait -PassThru
                        if ($installProcess.ExitCode -eq 0) {
                            Write-Host "    Installed $ext successfully" -ForegroundColor Green
                        } else {
                            Write-Host "    Installation failed for $ext (ExitCode: $($installProcess.ExitCode))" -ForegroundColor Yellow
                        }
                    }
                    catch {
                        Write-Host "    Error installing $ext - $_" -ForegroundColor Red
                    }
                }
            }
        }
    }
    Write-Host "`nSettings restoration complete!" -ForegroundColor Cyan
}
else {
    Write-Host "=== Backing up Editor Settings & Extensions ===" -ForegroundColor Cyan
    
    foreach ($config in $configs) {
        $name = $config.Name
        $userDir = $config.UserDir
        $repoDir = $config.RepoDir
        $cli = $config.CliCommand
        $isCmd = $config.IsCmd
        
        Write-Host ""
        Write-Host "Processing $name..." -ForegroundColor Cyan
        
        # Check if local installation user directory exists
        if (-not (Test-Path $userDir)) {
            Write-Host "Local user settings directory for $name does not exist. Skipping." -ForegroundColor Yellow
            continue
        }
        
        # Ensure repo folder for this profile exists
        if (-not (Test-Path $repoDir)) {
            New-Item -ItemType Directory -Path $repoDir -Force | Out-Null
        }
        
        # 1. Back up files
        foreach ($file in $filesToSync) {
            $sourcePath = Join-Path $userDir $file
            if (Test-Path $sourcePath) {
                $targetPath = Join-Path $repoDir $file
                Copy-Item -Path $sourcePath -Destination $targetPath -Recurse -Force
                Write-Host "  Backed up file/folder: $file" -ForegroundColor Green
            }
        }
        
        # 2. Back up/Export extension list
        $extensionsFile = Join-Path $repoDir "extensions.txt"
        Write-Host "  Exporting extension list..." -ForegroundColor DarkGray
        
        $cleanExtList = Get-Extensions -CliCommand $cli -IsCmd $isCmd
        
        if ($null -ne $cleanExtList) {
            if ($cleanExtList.Count -gt 0) {
                $cleanExtList | Out-File -FilePath $extensionsFile -Encoding utf8 -Force
                Write-Host "  Saved $($cleanExtList.Count) extensions to extensions.txt" -ForegroundColor Green
            } else {
                Write-Host "  No extensions found to export." -ForegroundColor Yellow
            }
        } else {
            Write-Host "  CLI for $name not found. Skipping extensions export." -ForegroundColor Yellow
        }
    }
    Write-Host "`nSettings backup complete!" -ForegroundColor Cyan
}

