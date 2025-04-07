<#
.SYNOPSIS
Active Directory user management using net user commands

.DESCRIPTION
Modifies user accounts with German localization and secure logging
#>

[CmdletBinding()]
param (
    [string]$LogFilePath = "C:\AD-Verwaltung\Protokolle.log",
    [int]$BatchSize = 500,
    [switch]$Silent
)

$ErrorActionPreference = "Stop"
$startTime = Get-Date

function Protokollieren {
    param ($Nachricht, $Level = "INFO")
    $Zeitstempel = Get-Date -Format "dd.MM.yyyy HH:mm:ss.fff"
    "$Zeitstempel [$Level] - $Nachricht" | Add-Content $LogFilePath -Encoding UTF8
    if (-not $Silent) { Write-Host "$Zeitstempel [$Level] - $Nachricht" }
}

try {
    Protokollieren "Skriptstart: NetUser-Implementierung"

    $users = Get-ADUser -Filter { 
        SamAccountName -like "L110*" -or SamAccountName -like "L114*" 
    } -Properties Enabled,PasswordNeverExpires -ResultPageSize $BatchSize |
    Where-Object { $_.SamAccountName -match '^(L110|L114)\d{4}$' }

    foreach ($user in $users) {
        try {
            $username = $user.SamAccountName
            $pass = "T3mp!" + (Get-Random -Minimum 100000 -Maximum 999999)
            
            # Account activation
            if (-not $user.Enabled) {
                net user $username /active:yes 2>&1 | Out-Null
                Protokollieren "Konto aktiviert: $username"
            }

            # Password change
            net user $username $pass 2>&1 | Out-Null
            Protokollieren "Passwort geändert: $username"

            # Password policies
            net user $username /passwordchg:no 2>&1 | Out-Null
            Set-ADUser $username -PasswordNeverExpires $true
            Protokollieren "Richtlinien aktualisiert: $username"

            # Force no password change at next logon
            Set-ADUser $username -Replace @{pwdLastSet = -1}
        }
        catch {
            Protokollieren "Fehler bei $username: $_" -Level "FEHLER"
            continue
        }
    }
}
catch {
    Protokollieren "Globaler Fehler: $_" -Level "FEHLER"
    exit 1
}
finally {
    $dauer = (Get-Date) - $startTime
    Protokollieren "Skriptdauer: $($dauer.TotalSeconds.ToString("N2")) Sekunden"
}
