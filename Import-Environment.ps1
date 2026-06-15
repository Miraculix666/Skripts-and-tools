<#
.SYNOPSIS
    Loads environmental configuration from a local environment.json file.

.DESCRIPTION
    This script reads and parses a local 'environment.json' configuration file,
    storing the result in the global variable '$global:Env'. This allows multi-OS scripts
    to access common configuration settings (such as Active Directory OUs, mail servers,
    output folders, etc.) across Windows, Linux, and macOS without hardcoding secrets.

.NOTES
    Version: 1.0
    Autor: Systems Administration
    Compatibility: PowerShell 5.1+
#>

$configFileName = "environment.json"

# Search for environment.json in the script's directory, then parent directories
$configPath = ""
$currentDir = $PSScriptRoot
if (-not $currentDir) {
    $currentDir = $PWD.Path
}

while ($currentDir) {
    $testPath = Join-Path $currentDir $configFileName
    if (Test-Path $testPath) {
        $configPath = $testPath
        break
    }
    # Go up one level
    $parent = Split-Path $currentDir -Parent
    if ($parent -eq $currentDir) {
        break
    }
    $currentDir = $parent
}

if ($configPath) {
    try {
        $jsonContent = Get-Content -Raw -Path $configPath -Encoding UTF8 -ErrorAction Stop
        $global:Env = ConvertFrom-Json $jsonContent -ErrorAction Stop
        Write-Verbose "Successfully loaded environment settings from: $configPath"
    } catch {
        Write-Error "Failed to load environment configuration from '$configPath': $_"
        $global:Env = $null
    }
} else {
    Write-Warning "Configuration file '$configFileName' not found in path hierarchy."
    $global:Env = $null
}
