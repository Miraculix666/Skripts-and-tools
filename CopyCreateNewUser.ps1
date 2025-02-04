# Advanced AD User Management Script
# Author: Bolt
# Version: 3.0
# Description: Creates new AD users based on template users with complete attribute copying
#requires -Module ActiveDirectory

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$TemplateUser,
    
    [Parameter(Mandatory=$false)]
    [string]$NewUserName,
    
    [Parameter(Mandatory=$false)]
    [string]$CSV,
    
    [Parameter(Mandatory=$false)]
    [switch]$ExportTemplate,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath
)

# Function to write verbose messages with timestamp
function Write-VerboseWithTime {
    param([string]$Message)
    Write-Verbose "$(Get-Date -Format 'dd.MM.yyyy HH:mm:ss'): $Message"
}

# Function to get all group memberships
function Get-UserGroupMemberships {
    param([string]$Username)
    
    Write-VerboseWithTime "Retrieving group memberships for user: $Username"
    try {
        $groups = Get-ADPrincipalGroupMembership -Identity $Username | Select-Object -ExpandProperty Name
        return $groups
    }
    catch {
        Write-Warning "Failed to get group memberships: $_"
        return @()
    }
}

# Function to get user's OU path
function Get-UserOUPath {
    param([string]$Username)
    
    Write-VerboseWithTime "Retrieving OU path for user: $Username"
    try {
        $user = Get-ADUser -Identity $Username -Properties DistinguishedName
        $ouPath = ($user.DistinguishedName -split ',', 2)[1]
        return $ouPath
    }
    catch {
        Write-Warning "Failed to get OU path: $_"
        return $null
    }
}

# Function to validate mandatory user properties
function Test-MandatoryProperties {
    param(
        [hashtable]$Properties
    )
    
    $mandatoryProps = @('SamAccountName', 'UserPrincipalName', 'GivenName', 'Surname')
    $missingProps = @()
    
    foreach ($prop in $mandatoryProps) {
        if (-not $Properties.ContainsKey($prop) -or [string]::IsNullOrWhiteSpace($Properties[$prop])) {
            $missingProps += $prop
        }
    }
    
    return $missingProps
}

# Function to get user input with validation
function Get-ValidatedInput {
    param(
        [string]$Prompt,
        [switch]$Required
    )
    
    do {
        $input = Read-Host -Prompt $Prompt
        if ($Required -and [string]::IsNullOrWhiteSpace($input)) {
            Write-Warning "Dieses Feld ist erforderlich. Bitte geben Sie einen Wert ein."
            $valid = $false
        } else {
            $valid = $true
        }
    } while (-not $valid)
    
    return $input
}

# Function to export user template to CSV
function Export-UserTemplate {
    param(
        [string]$Username,
        [string]$ExportPath
    )
    
    Write-VerboseWithTime "Exportiere Benutzervorlage '$Username' nach CSV"
    
    try {
        $user = Get-ADUser -Identity $Username -Properties *
        if (-not $user) {
            throw "Vorlagenbenutzer nicht gefunden"
        }
        
        # Get group memberships and OU path
        $groups = Get-UserGroupMemberships -Username $Username
        $ouPath = Get-UserOUPath -Username $Username
        
        $exportProperties = @(
            'SamAccountName',
            'UserPrincipalName',
            'GivenName',
            'Surname',
            'DisplayName',
            'Description',
            'Office',
            'Department',
            'Company',
            'Title',
            'Manager',
            'StreetAddress',
            'City',
            'State',
            'PostalCode',
            'Country',
            'HomePhone',
            'MobilePhone',
            'OfficePhone',
            'Fax',
            'EmailAddress',
            'ScriptPath',
            'HomeDrive',
            'HomeDirectory',
            'ProfilePath'
        )
        
        $userProps = [PSCustomObject]@{}
        foreach ($prop in $exportProperties) {
            $userProps | Add-Member -MemberType NoteProperty -Name $prop -Value $user.$prop
        }
        
        # Add group memberships and OU path
        $userProps | Add-Member -MemberType NoteProperty -Name 'GroupMemberships' -Value ($groups -join '|')
        $userProps | Add-Member -MemberType NoteProperty -Name 'OUPath' -Value $ouPath
        
        $fileName = if ([string]::IsNullOrWhiteSpace($ExportPath)) {
            "UserTemplate_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        } else {
            $ExportPath
        }
        
        $userProps | Export-Csv -Path $fileName -NoTypeInformation -Encoding UTF8 -Delimiter ";"
        Write-VerboseWithTime "Vorlage erfolgreich exportiert nach: $fileName"
    }
    catch {
        Write-Error "Fehler beim Exportieren der Vorlage: $_"
        return $false
    }
    
    return $true
}

# Function to create new user from template
function New-UserFromTemplate {
    param(
        [string]$TemplateUser,
        [hashtable]$NewUserProperties
    )
    
    Write-VerboseWithTime "Erstelle neuen Benutzer basierend auf Vorlage: $TemplateUser"
    
    try {
        # Get template user and OU path
        $template = Get-ADUser -Identity $TemplateUser -Properties *
        if (-not $template) {
            throw "Vorlagenbenutzer nicht gefunden"
        }
        
        $ouPath = Get-UserOUPath -Username $TemplateUser
        
        # Create new user properties
        $userProps = @{
            Instance = $template
            Enabled = $true
            ChangePasswordAtLogon = $true
            Path = $ouPath
        }
        
        # Add all new user properties
        foreach ($key in $NewUserProperties.Keys) {
            $userProps[$key] = $NewUserProperties[$key]
        }
        
        # Create the new user
        $newUser = New-ADUser @userProps -PassThru
        
        # Copy group memberships
        Write-VerboseWithTime "Kopiere Gruppenmitgliedschaften"
        $groups = Get-UserGroupMemberships -Username $TemplateUser
        foreach ($group in $groups) {
            try {
                Add-ADGroupMember -Identity $group -Members $newUser.SamAccountName
                Write-VerboseWithTime "Benutzer zur Gruppe '$group' hinzugefügt"
            }
            catch {
                Write-Warning "Konnte Benutzer nicht zur Gruppe '$group' hinzufügen: $_"
            }
        }
        
        Write-VerboseWithTime "Benutzer erfolgreich erstellt: $($NewUserProperties.SamAccountName)"
        return $true
    }
    catch {
        Write-Error "Fehler beim Erstellen des Benutzers: $_"
        return $false
    }
}

# Function to process CSV file
function Process-CSVFile {
    param(
        [string]$CSVPath,
        [string]$TemplateUser
    )
    
    Write-VerboseWithTime "Verarbeite CSV-Datei: $CSVPath"
    
    try {
        $users = Import-Csv -Path $CSVPath -Delimiter ";"
        foreach ($user in $users) {
            $userProps = @{}
            foreach ($prop in $user.PSObject.Properties) {
                if (-not [string]::IsNullOrWhiteSpace($prop.Value)) {
                    $userProps[$prop.Name] = $prop.Value
                }
            }
            
            $missingProps = Test-MandatoryProperties -Properties $userProps
            if ($missingProps.Count -gt 0) {
                Write-Warning "Fehlende Pflichtfelder für Benutzer: $($missingProps -join ', ')"
                continue
            }
            
            New-UserFromTemplate -TemplateUser $TemplateUser -NewUserProperties $userProps
        }
    }
    catch {
        Write-Error "Fehler beim Verarbeiten der CSV-Datei: $_"
        return $false
    }
    
    return $true
}

# Main script logic
try {
    # Set verbose output
    $VerbosePreference = "Continue"
    
    Write-VerboseWithTime "Starte AD-Benutzerverwaltungsskript"
    
    # Handle template export
    if ($ExportTemplate) {
        if ([string]::IsNullOrWhiteSpace($TemplateUser)) {
            $TemplateUser = Get-ValidatedInput -Prompt "Benutzername der Vorlage eingeben" -Required
        }
        Export-UserTemplate -Username $TemplateUser -ExportPath $ExportPath
        return
    }
    
    # Interactive mode if no parameters provided
    if (-not $TemplateUser -and -not $CSV) {
        Write-VerboseWithTime "Starte interaktiven Modus"
        $TemplateUser = Get-ValidatedInput -Prompt "Benutzername der Vorlage eingeben" -Required
        
        $mode = Get-ValidatedInput -Prompt "Modus wählen (single/csv)" -Required
        if ($mode -eq "csv") {
            $CSV = Get-ValidatedInput -Prompt "CSV-Dateipfad eingeben" -Required
        } else {
            $userProps = @{
                SamAccountName = Get-ValidatedInput -Prompt "Neuen Benutzernamen eingeben (SAMAccountName)" -Required
                UserPrincipalName = Get-ValidatedInput -Prompt "UserPrincipalName eingeben" -Required
                GivenName = Get-ValidatedInput -Prompt "Vorname eingeben" -Required
                Surname = Get-ValidatedInput -Prompt "Nachname eingeben" -Required
                DisplayName = Get-ValidatedInput -Prompt "Anzeigenamen eingeben"
                Description = Get-ValidatedInput -Prompt "Beschreibung eingeben"
            }
            
            New-UserFromTemplate -TemplateUser $TemplateUser -NewUserProperties $userProps
        }
    }
    # CSV mode
    elseif ($CSV) {
        if (-not $TemplateUser) {
            $TemplateUser = Get-ValidatedInput -Prompt "Benutzername der Vorlage eingeben" -Required
        }
        Process-CSVFile -CSVPath $CSV -TemplateUser $TemplateUser
    }
    # Parameter mode
    else {
        if (-not $NewUserName) {
            $NewUserName = Get-ValidatedInput -Prompt "Neuen Benutzernamen eingeben (SAMAccountName)" -Required
        }
        
        $userProps = @{
            SamAccountName = $NewUserName
            UserPrincipalName = Get-ValidatedInput -Prompt "UserPrincipalName eingeben" -Required
            GivenName = Get-ValidatedInput -Prompt "Vorname eingeben" -Required
            Surname = Get-ValidatedInput -Prompt "Nachname eingeben" -Required
            DisplayName = Get-ValidatedInput -Prompt "Anzeigenamen eingeben"
            Description = Get-ValidatedInput -Prompt "Beschreibung eingeben"
        }
        
        New-UserFromTemplate -TemplateUser $TemplateUser -NewUserProperties $userProps
    }
    
    Write-VerboseWithTime "Skript erfolgreich beendet"
}
catch {
    Write-Error "Skriptfehler: $_"
    Write-VerboseWithTime "Skript mit Fehlern beendet"
}
finally {
    $VerbosePreference = "SilentlyContinue"
}
