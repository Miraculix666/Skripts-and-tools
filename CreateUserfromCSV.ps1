# Active Directory User Management Script
# Author: Enhanced by Bolt
# Version: 2.4
# PowerShell Version: 5.1
# Description:
#   - Creates new AD users from a CSV file
#   - Assigns groups based on CSV data
#   - Supports comprehensive logging and verbose output
#   - Implements German localization

[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory = $false, HelpMessage = "Path to CSV file for user import")]
    [string]$CsvPath = "neue_benutzer.csv",
    
    [Parameter(Mandatory = $false, HelpMessage = "Default password for new users")]
    [SecureString]$DefaultPassword = (ConvertTo-SecureString "Willkommen2024!" -AsPlainText -Force),
    
    [Parameter(Mandatory = $false, HelpMessage = "Default OU for new users")]
    [string]$DefaultOU,
    
    [Parameter(Mandatory = $false, HelpMessage = "Log file path")]
    [string]$LogPath = "ADBenutzerVerwaltung.log",
    
    [Parameter(Mandatory = $false, HelpMessage = "Enable verbose output")]
    [switch]$Verbose
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
        [Parameter(Mandatory = $false)]
        [string]$GivenName,
        [Parameter(Mandatory = $false)]
        [string]$Surname,
        [Parameter(Mandatory = $false)]
        [string]$Department,
        [Parameter(Mandatory = $false)]
        [string]$Title,
        [Parameter(Mandatory = $false)]
        [string]$Manager,
        [Parameter(Mandatory = $false)]
        [string]$Office,
        [Parameter(Mandatory = $false)]
        [string]$OfficePhone,
        [Parameter(Mandatory = $false)]
        [string]$Company,
        [Parameter(Mandatory = $false)]
        [string]$Description,
        [Parameter(Mandatory = $true)]
        [string]$OU,
        [string[]]$Groups,
        [SecureString]$Password = $DefaultPassword
    )
    
    Write-Verbose "Erstelle neuen Benutzer: $Name in OU: $OU"
    
    try {
        if ($PSCmdlet.ShouldProcess($Name, "Benutzer erstellen")) {
            # Validate OU exists
            if (-not (Get-ADOrganizationalUnit -Filter {DistinguishedName -eq $OU})) {
                throw "Die angegebene OU existiert nicht: $OU"
            }
            
            $userParams = @{
                Name = $Name
                SamAccountName = $SamAccountName
                UserPrincipalName = $UserPrincipalName
                Path = $OU
                AccountPassword = $Password
                Enabled = $true
                ChangePasswordAtLogon = $false # Kein Passwortwechsel erforderlich
            }
            
            # Add optional parameters if provided
            if ($GivenName) { $userParams.GivenName = $GivenName }
            if ($Surname) { $userParams.Surname = $Surname }
            if ($Department) { $userParams.Department = $Department }
            if ($Title) { $userParams.Title = $Title }
            if ($Manager) { $userParams.Manager = $Manager }
            if ($Office) { $userParams.Office = $Office }
            if ($OfficePhone) { $userParams.OfficePhone = $OfficePhone }
            if ($Company) { $userParams.Company = $Company }
            if ($Description) { $userParams.Description = $Description }
            
            New-ADUser @userParams
            
            foreach ($group in $Groups) {
                Add-ADGroupMember -Identity $group -Members $SamAccountName
                Write-Verbose "Gruppe '$group' dem Benutzer '$SamAccountName' zugewiesen"
            }
            
            Write-CustomLog "Benutzer '$Name' erfolgreich erstellt in OU: $OU" -Level "INFO"
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
            # Determine OU based on priority: CSV > DefaultOU
            $ou = if ($user.OU) { 
                $user.OU 
            } elseif ($DefaultOU) { 
                $DefaultOU 
            } else { 
                Write-CustomLog "Keine OU für Benutzer '$($user.Benutzername)' angegeben. Überspringe..." -Level "WARNUNG"
                continue
            }
            
            if ($PSCmdlet.ShouldProcess($user.Name, "Benutzer aus CSV erstellen")) {
                $params = @{
                    SamAccountName = $user.Benutzername
                    UserPrincipalName = $user.EMail
                    Name = if ($user.'Vollständiger Name') { $user.'Vollständiger Name' } else { "$($user.Vorname) $($user.Nachname)" }
                    OU = $ou
                    Groups = if ($user.Gruppen) { ($user.Gruppen -split ';') } else { @() }
                }
                
                # Add optional parameters if they exist in CSV
                if ($user.Vorname) { $params.GivenName = $user.Vorname }
                if ($user.Nachname) { $params.Surname = $user.Nachname }
                if ($user.Abteilung) { $params.Department = $user.Abteilung }
                if ($user.Position) { $params.Title = $user.Position }
                if ($user.Vorgesetzter) { $params.Manager = $user.Vorgesetzter }
                if ($user.Büro) { $params.Office = $user.Büro }
                if ($user.Telefon) { $params.OfficePhone = $user.Telefon }
                if ($user.Firma) { $params.Company = $user.Firma }
                if ($user.Beschreibung) { $params.Description = $user.Beschreibung }
                
                New-CustomADUser @params
            }
        }
    } catch {
        Write-CustomLog "Fehler beim Import aus CSV: $_" -Level "FEHLER"
    }
}

# Main execution block
try {
    Write-CustomLog "Skript-Ausführung gestartet" -Level "INFO"
    
    Import-ADUsersFromCSV
    
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
    Dieses Skript ermöglicht das Erstellen neuer AD-Benutzer basierend auf einem CSV-File.

PARAMETER:
    -CsvPath         Pfad zur Import-CSV-Datei
    -DefaultOU       Standard-OU für neue Benutzer
    -DefaultPassword Standardpasswort für neue Benutzer
    -Verbose         Aktiviert detaillierte Ausgabe
    -WhatIf          Zeigt an, was passieren würde, ohne Änderungen vorzunehmen

BEISPIELE:
    .\ADUserCreation.ps1 -CsvPath "neue_benutzer.csv" -Verbose
    .\ADUserCreation.ps1 -CsvPath "neue_benutzer.csv" -WhatIf
"@
