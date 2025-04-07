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
