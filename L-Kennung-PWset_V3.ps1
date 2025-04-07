<#
.SYNOPSIS
Modifies user accounts without checking existing properties

.DESCRIPTION
Applies settings to all users matching naming patterns
#>

[CmdletBinding()]
param (
    [string]$LogPath = "C:\ADLogs\Operations.log",
    [int]$BatchSize = 500,
    [switch]$Silent
)

$ErrorActionPreference = "Stop"
$startTime = Get-Date

function Write-Log {
    param($Message, $Level = "INFO")
    $timestamp = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
    "$timestamp [$Level] - $Message" | Add-Content $LogPath -Encoding UTF8
    if (-not $Silent) { Write-Host "$timestamp [$Level] - $Message" }
}

try {
    Write-Log "Starting user modification process"

    # Get users without property checks
    $users = Get-ADUser -Filter {
        SamAccountName -like "L110*" -or SamAccountName -like "L114*"
    } -ResultPageSize $BatchSize |
    Where-Object { $_.SamAccountName -match '^(L110|L114)\d{4}$' }

    Write-Log "Found $($users.Count) users for processing"

    foreach ($user in $users) {
        try {
            $username = $user.SamAccountName
            $tempPass = "TempPass$(Get-Random -Min 1000 -Max 9999)!"

            # Always reactivate account
            net user $username /active:yes 2>&1 | Out-Null
            Write-Log "Reactivated account: $username"

            # Set new password
            net user $username $tempPass 2>&1 | Out-Null
            Write-Log "Password reset for: $username"

            # Apply password policies
            net user $username /passwordchg:no 2>&1 | Out-Null
            Set-ADUser $username -PasswordNeverExpires $true
            Set-ADUser $username -Replace @{pwdLastSet = -1}
            Write-Log "Policy applied: $username"
        }
        catch {
            Write-Log "Error processing $username : $_" -Level "ERROR"
            continue
        }
    }
}
catch {
    Write-Log "Critical error: $_" -Level "ERROR"
    exit 1
}
finally {
    $duration = (Get-Date) - $startTime
    Write-Log "Operation completed in $($duration.TotalSeconds.ToString('N2')) seconds"
}
