# Active Directory User Management Script
# Author: Enhanced by Bolt
# Version: 2.0
# PowerShell Version: 5.1
# Description:
# - Exports existing AD users to CSV based on a template user
# - Creates new users interactively, via parameters, or from CSV
# - Assigns groups based on template user
# - Supports comprehensive logging and verbose output
# - Implements German localization

[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory = $false, HelpMessage = "Template user for group assignments")]
    [string]$TemplateUser,
    
    [Parameter(Mandatory = $false, HelpMessage = "Path for exporting user data (CSV)")]
    [string]$ExportPath = "ADBenutzerExport.csv",
    
    [Parameter(Mandatory = $false, HelpMessage = "Path to CSV file for user import")]
    [string]$CsvPath,
    
    [Parameter(Mandatory = $false, HelpMessage = "Log file path")]
    [string]$LogPath = "ADBenutzerVerwaltung.log",
    
    [Parameter(Mandatory = $false, HelpMessage = "Default password for new users")]
    [SecureString]$DefaultPassword = (ConvertTo-SecureString "Willkommen2024!" -AsPlainText -Force)
)

# Set German culture for proper date/number formatting
$previousCulture = [System.Threading.Thread]::CurrentThread.CurrentCulture
[System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo("de-DE")

# Import required modules with error handling
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Verbose "ActiveDirectory-Modul erfolgreich geladen"
} catch {
    Write-Error "Fehler beim Laden des ActiveDirectory-Moduls: $_"
    exit 1
}

# Enhanced logging function with severity levels
function Write-CustomLog {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNUNG", "FEHLER", "DEBUG")]
        [string]$Level = "INFO"
    )
    
    $Timestamp = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
    $LogMessage = "$Timestamp [$Level] $Message"
    
    # Write to log file
    Add-Content -Path $LogPath -Value $LogMessage -Encoding UTF8
    
    # Console output with color coding
    $Color = switch ($Level) {
        "INFO"    { "White" }
        "WARNUNG" { "Yellow" }
        "FEHLER"  { "Red" }
        "DEBUG"   { "Cyan" }
    }
    Write-Host $LogMessage -ForegroundColor $Color
    
    # Additional verbose output for debugging
    if ($Level -eq "DEBUG") {
        Write-Verbose $Message
    }
}

# Function to validate and get AD user
function Get-ValidADUser {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Identity,
        [string]$Operation = "Abfrage"
    )
    
    try {
        $user = Get-ADUser -Identity $Identity -Properties MemberOf, DistinguishedName -ErrorAction Stop
        Write-Verbose "Benutzer '$Identity' erfolgreich gefunden"
        return $user
    } catch {
        Write-CustomLog "Fehler bei $Operation für Benutzer '$Identity': $_" -Level "FEHLER"
        return $null
    }
}

# Function to get template user's OU
function Get-TemplateUserOU {
    param (
        [Parameter(Mandatory = $true)]
        [string]$TemplateUser
    )
    
    $user = Get-ValidADUser -Identity $TemplateUser -Operation "OU-Ermittlung"
    if ($user) {
        return ($user.DistinguishedName -split ",", 2)[1]
    }
    return $null
}

# Function to export AD users
function Export-ADUsers {
    [CmdletBinding()]
    param()
    
    Write-CustomLog "Starte Benutzerexport basierend auf Template: $TemplateUser" -Level "INFO"
    
    $Template = Get-ValidADUser -Identity $TemplateUser -Operation "Template-Validierung"
    if (-not $Template) { return }
    
    try {
        $users = Get-ADUser -Filter * -Properties SamAccountName, UserPrincipalName, Name, MemberOf |
                Where-Object { $_.MemberOf -contains $Template.MemberOf }
        
        $exportData = $users | Select-Object @{
            Name='Benutzername'; Expression={$_.SamAccountName}
        }, @{
            Name='E-Mail'; Expression={$_.UserPrincipalName}
        }, @{
            Name='Vollständiger Name'; Expression={$_.Name}
        }, @{
            Name='Gruppen'; Expression={$_.MemberOf -join ';'}
        }
        
        $exportData | Export-Csv -Path $ExportPath -Delimiter ";" -NoTypeInformation -Encoding UTF8
        Write-CustomLog "Benutzerdaten erfolgreich exportiert nach: $ExportPath" -Level "INFO"
    } catch {
        Write-CustomLog "Fehler beim Exportieren der Benutzerdaten: $_" -Level "FEHLER"
    }
}

# Function to create new AD user
function New-CustomADUser {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true)]
        [string]$SamAccountName,
        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName,
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        [string]$OU,
        [string[]]$Groups,
        [SecureString]$Password = $DefaultPassword
    )
    
    Write-Verbose "Erstelle neuen Benutzer: $Name"
    
    try {
        if ($PSCmdlet.ShouldProcess($Name, "Benutzer erstellen")) {
            $userParams = @{
                Name = $Name
                SamAccountName = $SamAccountName
                UserPrincipalName = $UserPrincipalName
                Path = $OU
                AccountPassword = $Password
                Enabled = $true
            }
            
            New-ADUser @userParams
            
            foreach ($group in $Groups) {
                Add-ADGroupMember -Identity $group -Members $SamAccountName
                Write-Verbose "Gruppe '$group' dem Benutzer '$SamAccountName' zugewiesen"
            }
            
            Write-CustomLog "Benutzer '$Name' erfolgreich erstellt" -Level "INFO"
        }
    } catch {
        Write-CustomLog "Fehler beim Erstellen von Benutzer '$Name': $_" -Level "FEHLER"
    }
}

# Function to create users from CSV
function Import-ADUsersFromCSV {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    
    if (-not (Test-Path $CsvPath)) {
        Write-CustomLog "CSV-Datei nicht gefunden: $CsvPath" -Level "FEHLER"
        return
    }
    
    Write-CustomLog "Importiere Benutzer aus CSV: $CsvPath" -Level "INFO"
    
    try {
        $users = Import-Csv -Path $CsvPath -Delimiter ";" -Encoding UTF8
        foreach ($user in $users) {
            $ou = if ($user.OU) { $user.OU } else { Get-TemplateUserOU -TemplateUser $TemplateUser }
            
            if ($PSCmdlet.ShouldProcess($user.Name, "Benutzer aus CSV erstellen")) {
                New-CustomADUser -SamAccountName $user.Benutzername `
                               -UserPrincipalName $user.EMail `
                               -Name $user.'Vollständiger Name' `
                               -OU $ou `
                               -Groups ($user.Gruppen -split ';')
            }
        }
    } catch {
        Write-CustomLog "Fehler beim Import aus CSV: $_" -Level "FEHLER"
    }
}

# Main execution block
try {
    Write-CustomLog "Skript-Ausführung gestartet" -Level "INFO"
    
    # Interactive parameter collection if needed
    if (-not $TemplateUser) {
        $TemplateUser = Read-Host "Bitte Template-Benutzer eingeben"
    }
    
    if ($TemplateUser) {
        Export-ADUsers
    }
    
    if ($CsvPath) {
        Import-ADUsersFromCSV
    }
    
    Write-CustomLog "Skript-Ausführung erfolgreich beendet" -Level "INFO"
} catch {
    Write-CustomLog "Unerwarteter Fehler bei der Skript-Ausführung: $_" -Level "FEHLER"
} finally {
    # Restore original culture
    [System.Threading.Thread]::CurrentThread.CurrentCulture = $previousCulture
}

# Help message
$HelpMessage = @"
Active Directory Benutzerverwaltungsskript
----------------------------------------

BESCHREIBUNG:
    Dieses Skript ermöglicht das Exportieren von AD-Benutzern und das Erstellen neuer Benutzer
    basierend auf einem Template-Benutzer.

PARAMETER:
    -TemplateUser     Template-Benutzer für Gruppenberechtigungen
    -ExportPath       Pfad für CSV-Export (Standard: ADBenutzerExport.csv)
    -CsvPath         Pfad zur Import-CSV-Datei
    -LogPath         Pfad zur Log-Datei (Standard: ADBenutzerVerwaltung.log)
    -Verbose         Aktiviert detaillierte Ausgabe
    -WhatIf          Zeigt an, was passieren würde, ohne Änderungen vorzunehmen

BEISPIELE:
    .\ADUserManagement.ps1 -TemplateUser "mustermax" -Verbose
    .\ADUserManagement.ps1 -CsvPath "neue_benutzer.csv" -WhatIf
"@

if ($MyInvocation.Line -match '-(\?|Help)') {
    Write-Host $HelpMessage
    exit 0
}
