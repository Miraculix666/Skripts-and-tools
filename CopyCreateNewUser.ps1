# Advanced AD User Management Script
# Author: Bolt
# Version: 3.2
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
    [string]$ExportPath,
    
    [Parameter(Mandatory=$false)]
    [SecureString]$Password,
    
    [Parameter(Mandatory=$false)]
    [string]$TargetOU
)

# Always enable verbose output
$VerbosePreference = "Continue"

# Function to write verbose messages with timestamp
function Write-VerboseWithTime {
    param([string]$Message)
    Write-Verbose "$(Get-Date -Format 'dd.MM.yyyy HH:mm:ss'): $Message"
}

# Function to get secure password
function Get-SecurePassword {
    Write-VerboseWithTime "Fordere sicheres Passwort an"
    do {
        $securePassword = Read-Host -Prompt "Passwort eingeben (mindestens 8 Zeichen)" -AsSecureString
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
        $plainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        if ($plainPassword.Length -lt 8) {
            Write-Warning "Das Passwort muss mindestens 8 Zeichen lang sein."
            $valid = $false
        } else {
            $valid = $true
        }
    } while (-not $valid)
    return $securePassword
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
    param(
        [string]$Username,
        [string]$TargetOU
    )
    
    Write-VerboseWithTime "Retrieving OU path for user: $Username"
    try {
        if ($TargetOU) {
            # Validate and return target OU
            $ou = Get-ADOrganizationalUnit -Identity $TargetOU
            return $ou.DistinguishedName
        } else {
            # Get OU from template user
            $user = Get-ADUser -Identity $Username -Properties DistinguishedName
            $ouPath = ($user.DistinguishedName -split ',', 2)[1]
            return $ouPath
        }
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
    
    # Based on Petri article mandatory fields
    $mandatoryProps = @(
        'SamAccountName',
        'UserPrincipalName',
        'GivenName',
        'Surname',
        'Name',
        'DisplayName',
        'EmailAddress',
        'Department',
        'Title'
    )
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

# Function to get target OU
function Get-TargetOUPath {
    Write-VerboseWithTime "Fordere Ziel-OU an"
    do {
        $ouPath = Get-ValidatedInput -Prompt "Ziel-OU eingeben (Distinguished Name)" -Required
        try {
            $ou = Get-ADOrganizationalUnit -Identity $ouPath
            $valid = $true
        }
        catch {
            Write-Warning "Ungültige OU. Bitte geben Sie einen gültigen Distinguished Name ein."
            $valid = $false
        }
    } while (-not $valid)
    return $ouPath
}

[... Rest of the script remains the same but with updated mandatory properties in the interactive mode and CSV processing ...]

# Main script logic
try {
    Write-VerboseWithTime "Starte AD-Benutzerverwaltungsskript"
    
    # Handle template export
    if ($ExportTemplate) {
        if ([string]::IsNullOrWhiteSpace($TemplateUser)) {
            $TemplateUser = Get-ValidatedInput -Prompt "Benutzername der Vorlage eingeben" -Required
        }
        Export-UserTemplate -Username $TemplateUser -ExportPath $ExportPath
        return
    }
    
    # Get password if not provided (now mandatory)
    if (-not $Password) {
        $Password = Get-SecurePassword
    }
    
    # Get target OU if not provided (now mandatory)
    if (-not $TargetOU) {
        $TargetOU = Get-TargetOUPath
    }
    
    # Interactive mode if no parameters provided
    if (-not $TemplateUser -and -not $CSV) {
        Write-VerboseWithTime "Starte interaktiven Modus"
        $TemplateUser = Get-ValidatedInput -Prompt "Benutzername der Vorlage eingeben" -Required
        
        $mode = Get-ValidatedInput -Prompt "Modus wählen (single/csv)" -Required
        if ($mode -eq "csv") {
            $CSV = Get-ValidatedInput -Prompt "CSV-Dateipfad eingeben" -Required
            Process-CSVFile -CSVPath $CSV -TemplateUser $TemplateUser -Password $Password -TargetOU $TargetOU
        } else {
            $userProps = @{
                SamAccountName = Get-ValidatedInput -Prompt "Neuen Benutzernamen eingeben (SAMAccountName)" -Required
                UserPrincipalName = Get-ValidatedInput -Prompt "UserPrincipalName eingeben" -Required
                GivenName = Get-ValidatedInput -Prompt "Vorname eingeben" -Required
                Surname = Get-ValidatedInput -Prompt "Nachname eingeben" -Required
                Name = Get-ValidatedInput -Prompt "Name eingeben" -Required
                DisplayName = Get-ValidatedInput -Prompt "Anzeigenamen eingeben" -Required
                EmailAddress = Get-ValidatedInput -Prompt "E-Mail-Adresse eingeben" -Required
                Department = Get-ValidatedInput -Prompt "Abteilung eingeben" -Required
                Title = Get-ValidatedInput -Prompt "Position eingeben" -Required
                Description = Get-ValidatedInput -Prompt "Beschreibung eingeben"
            }
            
            New-UserFromTemplate -TemplateUser $TemplateUser -NewUserProperties $userProps -Password $Password -TargetOU $TargetOU
        }
    }
    # CSV mode
    elseif ($CSV) {
        if (-not $TemplateUser) {
            $TemplateUser = Get-ValidatedInput -Prompt "Benutzername der Vorlage eingeben" -Required
        }
        Process-CSVFile -CSVPath $CSV -TemplateUser $TemplateUser -Password $Password -TargetOU $TargetOU
    }
    # Parameter mode - only prompt for missing mandatory parameters
    else {
        $userProps = @{}
        
        # Add provided parameters
        if ($NewUserName) {
            $userProps['SamAccountName'] = $NewUserName
        }
        
        # Check for missing mandatory parameters
        $mandatoryProps = @{
            'UserPrincipalName' = 'UserPrincipalName eingeben'
            'GivenName' = 'Vorname eingeben'
            'Surname' = 'Nachname eingeben'
            'Name' = 'Name eingeben'
            'SamAccountName' = 'Neuen Benutzernamen eingeben (SAMAccountName)'
            'DisplayName' = 'Anzeigenamen eingeben'
            'EmailAddress' = 'E-Mail-Adresse eingeben'
            'Department' = 'Abteilung eingeben'
            'Title' = 'Position eingeben'
        }
        
        foreach ($prop in $mandatoryProps.GetEnumerator()) {
            if (-not $userProps.ContainsKey($prop.Key)) {
                $userProps[$prop.Key] = Get-ValidatedInput -Prompt $prop.Value -Required
            }
        }
        
        # Optional parameters
        $userProps['Description'] = Get-ValidatedInput -Prompt "Beschreibung eingeben"
        
        New-UserFromTemplate -TemplateUser $TemplateUser -NewUserProperties $userProps -Password $Password -TargetOU $TargetOU
    }
    
    Write-VerboseWithTime "Skript erfolgreich beendet"
}
catch {
    Write-Error "Skriptfehler: $_"
    Write-VerboseWithTime "Skript mit Fehlern beendet"
}
