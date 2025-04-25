<#
.SYNOPSIS
Modifies user accounts in Active Directory using "net user" commands.

.DESCRIPTION
This script applies settings to user accounts matching specific naming patterns (e.g., "L110####" or "L114####"). 
It ensures:
- Accounts are reactivated
- Passwords are reset
- Password policies are applied

Includes German localization and detailed logging.

.PARAMETER LogPath
Path to the log file where operations will be recorded.

.PARAMETER BatchSize
Number of users to process per batch.

.PARAMETER Silent
Suppresses console output if specified.

.EXAMPLE
.\Set-UserProperties.ps1 -LogPath "C:\Logs\UserUpdates.log" -BatchSize 500

.EXAMPLE
.\Set-UserProperties.ps1 -LogPath "C:\Logs\UserUpdates.log" -Silent

.NOTES
Requires administrative privileges.
#>

[CmdletBinding()]
param (
    [string]$LogPath = "C:\ADLogs\Operations.log",
    [int]$BatchSize = 500,
    [switch]$Silent
)

# Set error handling and localization
$ErrorActionPreference = "Stop"
$startTime = Get-Date

# Logging function with automatic directory creation
function Write-Log {
    param (
        [string]$Message,
        [ValidateSet("INFO", "WARNUNG", "FEHLER")]
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "dd.MM.yyyy HH:mm:ss.fff"
    $logEntry = "$timestamp [$Level] - $Message"

    # Ensure directory exists for the log file
    $logDirectory = Split-Path $LogPath -Parent
    if (-not (Test-Path -Path $logDirectory)) {
        New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
    }

    # Write log entry to file
    Add-Content -Path $LogPath -Value $logEntry -Encoding UTF8

    # Output to console unless Silent mode is enabled
    if (-not $Silent) {
        $colors = @{"INFO"="White"; "WARNUNG"="Yellow"; "FEHLER"="Red"}
        Write-Host $logEntry -ForegroundColor $colors[$Level]
    }
}

try {
    Write-Log "Script started: User modification process"

    # Retrieve users matching the naming pattern without checking properties
    $users = Get-ADUser -Filter {
        SamAccountName -like "L110*" -or SamAccountName -like "L114*"
    } -ResultPageSize $BatchSize |
    Where-Object { $_.SamAccountName -match '^(L110|L114)\d{4}$' }

    Write-Log "Found $($users.Count) users for processing"

    foreach ($user in $users) {
        try {
            $username = $user.SamAccountName
            $tempPass = "TempPass$(Get-Random -Min 1000 -Max 9999)!"

            # Reactivate account (always)
            net user $username /active:yes 2>&1 | Out-Null
            Write-Log "Account reactivated: $username"

            # Set new password using net user
            net user $username $tempPass 2>&1 | Out-Null
            Write-Log "Password reset for: $username"

            # Apply password policies using net user and AD cmdlets
            net user $username /passwordchg:no 2>&1 | Out-Null
            Set-ADUser -Identity $username -PasswordNeverExpires $true 2>&1 | Out-Null
            Set-ADUser -Identity $username -Replace @{pwdLastSet = 0} 2>&1 | Out-Null

            Write-Log "Password policies applied for: $username"
        }
        catch {
            Write-Log "Error processing user '$username': $_" -Level "FEHLER"
            continue
        }
    }
}
catch {
    Write-Log "Critical error encountered: $_" -Level "FEHLER"
    exit 1
}
finally {
    # Log script completion time and duration
    $duration = (Get-Date) - $startTime
    Write-Log "Script completed in $($duration.TotalSeconds.ToString('N2')) seconds"
}
