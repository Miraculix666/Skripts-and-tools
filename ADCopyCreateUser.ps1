# Active Directory User Management Script
# Author: Enhanced by Bolt
# Version: 2.2
# PowerShell Version: 5.1
# Description:
# - Exports existing AD users to CSV based on a template user's OUs
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
    [SecureString]$DefaultPassword = (ConvertTo-SecureString "Willkommen2024!" -AsPlainText -Force),
    
    [Parameter(Mandatory = $false, HelpMessage = "Default OU for new users")]
    [string]$DefaultOU
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

# Function to validate and get AD user with extended properties
function Get-ValidADUser {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Identity,
        [string]$Operation = "Abfrage"
    )
    
    try {
        $user = Get-ADUser -Identity $Identity -Properties MemberOf, DistinguishedName, GivenName, Surname, Department, Title, Manager, Office, OfficePhone, EmailAddress, Company, Description -ErrorAction Stop
        Write-Verbose "Benutzer '$Identity' erfolgreich gefunden"
        return $user
    } catch {
        Write-CustomLog "Fehler bei $Operation für Benutzer '$Identity': $_" -Level "FEHLER"
        return $null
    }
}

# Function to get all OUs from template user's path
function Get-TemplateUserOUs {
    param (
        [Parameter(Mandatory = $true)]
        [string]$TemplateUser
    )
    
    $user = Get-ValidADUser -Identity $TemplateUser -Operation "OU-Ermittlung"
    if ($user) {
        $ouPath = ($user.DistinguishedName -split ",", 2)[1]
        $ous = @()
        
        # Get all OUs from the path
        try {
            $ous = Get-ADOrganizationalUnit -Filter * -SearchBase $ouPath -SearchScope Subtree |
                   Select-Object -ExpandProperty DistinguishedName
            
            # Add the template user's direct OU
            $ous += $ouPath
            
            Write-Verbose "Gefundene OUs: $($ous.Count)"
            return $ous | Sort-Object -Unique
        } catch {
            Write-CustomLog "Fehler beim Abrufen der OUs: $_" -Level "FEHLER"
            return @($ouPath)
        }
    }
    return $null
}

# Function to compare group memberships
function Compare-GroupMembership {
    param (
        [string[]]$TemplateGroups,
        [string[]]$UserGroups
    )
    
    $templateGroupNames = $TemplateGroups | ForEach-Object { ($_ -split ',')[0] }
    $userGroupNames = $UserGroups | ForEach-Object { ($_ -split ',')[0] }
    
    $commonGroups = Compare-Object -ReferenceObject $templateGroupNames -DifferenceObject $userGroupNames -IncludeEqual -ExcludeDifferent
    return $commonGroups.Count -gt 0
}

# Function to export AD users
function Export-ADUsers {
    [CmdletBinding()]
    param()
    
    Write-CustomLog "Starte Benutzerexport basierend auf Template: $TemplateUser" -Level "INFO"
    
    $Template = Get-ValidADUser -Identity $TemplateUser -Operation "Template-Validierung"
    if (-not $Template) { return }
    
    try {
        Write-Verbose "Hole Template-Benutzer Gruppen und OUs"
        $templateGroups = $Template.MemberOf
        $templateOUs = Get-TemplateUserOUs -TemplateUser $TemplateUser
        
        Write-Verbose "Suche Benutzer in allen relevanten OUs"
        $allUsers = @()
        foreach ($ou in $templateOUs) {
            Write-Verbose "Durchsuche OU: $ou"
            $usersInOU = Get-ADUser -Filter * -SearchBase $ou -SearchScope OneLevel -Properties SamAccountName, UserPrincipalName, Name, GivenName, Surname, Department, Title, Manager, Office, OfficePhone, EmailAddress, Company, Description, MemberOf |
                        Where-Object { Compare-GroupMembership -TemplateGroups $templateGroups -UserGroups $_.MemberOf }
            $allUsers += $usersInOU
        }
        
        Write-Verbose "Gefundene Benutzer: $($allUsers.Count)"
        
        $exportData = $allUsers | Select-Object @{
            Name='Benutzername'; Expression={$_.SamAccountName}
        }, @{
            Name='E-Mail'; Expression={$_.EmailAddress}
        }, @{
            Name='Vorname'; Expression={$_.GivenName}
        }, @{
            Name='Nachname'; Expression={$_.Surname}
        }, @{
            Name='Abteilung'; Expression={$_.Department}
        }, @{
            Name='Position'; Expression={$_.Title}
        }, @{
            Name='Vorgesetzter'; Expression={$_.Manager}
        }, @{
            Name='Büro'; Expression={$_.Office}
        }, @{
            Name='Telefon'; Expression={$_.OfficePhone}
        }, @{
            Name='Firma'; Expression={$_.Company}
        }, @{
            Name='Beschreibung'; Expression={$_.Description}
        }, @{
            Name='OU'; Expression={($_.DistinguishedName -split ',',2)[1]}
        }, @{
            Name='Gruppen'; Expression={$_.MemberOf -join ';'}
        }
        
        $exportData | Export-Csv -Path $ExportPath -Delimiter ";" -NoTypeInformation -Encoding UTF8
        Write-CustomLog "Benutzerdaten erfolgreich exportiert nach: $ExportPath" -Level "INFO"
        Write-CustomLog "Anzahl exportierter Benutzer: $($exportData.Count)" -Level "INFO"
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
                ChangePasswordAtLogon = $true
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
        $templateOUs = Get-TemplateUserOUs -TemplateUser $TemplateUser
        $users = Import-Csv -Path $CsvPath -Delimiter ";" -Encoding UTF8
        
        foreach ($user in $users) {
            # Determine OU based on priority: CSV > DefaultOU > Random Template OU
            $ou = if ($user.OU -and ($templateOUs -contains $user.OU)) { 
                $user.OU 
            } elseif ($DefaultOU) { 
                $DefaultOU 
            } else { 
                $templateOUs | Get-Random
            }
            
            if ($PSCmdlet.ShouldProcess($user.Name, "Benutzer aus CSV erstellen")) {
                $params = @{
                    SamAccountName = $user.Benutzername
                    UserPrincipalName = $user.EMail
                    Name = $user.'Vollständiger Name'
                    OU = $ou
                    Groups = ($user.Gruppen -split ';')
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
    basierend auf einem Template-Benutzer. Es berücksichtigt alle OUs des Template-Benutzers.

PARAMETER:
    -TemplateUser     Template-Benutzer für Gruppenberechtigungen
    -ExportPath       Pfad für CSV-Export (Standard: ADBenutzerExport.csv)
    -CsvPath         Pfad zur Import-CSV-Datei
    -LogPath         Pfad zur Log-Datei (Standard: ADBenutzerVerwaltung.log)
    -DefaultOU       Standard-OU für neue Benutzer
    -Verbose         Aktiviert detaillierte Ausgabe
    -WhatIf          Zeigt an, was passieren würde, ohne Änderungen vorzunehmen

BEISPIELE:
    .\ADUserManagement.ps1 -TemplateUser "mustermax" -Verbose
    .\ADUserManagement.ps1 -CsvPath "neue_benutzer.csv" -WhatIf
    .\ADUserManagement.ps1 -TemplateUser "mustermax" -DefaultOU "OU=Users,DC=domain,DC=com"
"@

if ($MyInvocation.Line -match '-(\?|Help)') {
    Write-Host $HelpMessage
    exit 0
}
