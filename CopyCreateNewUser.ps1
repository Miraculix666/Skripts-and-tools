# Advanced AD User Management Script
# Author: Bolt
# Version: 3.3
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
    Write-VerboseWithTime "Fordere Passwort an"
    $securePassword = Read-Host -Prompt "Passwort eingeben" -AsSecureString
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

# Function to export user template to CSV
function Export-UserTemplate {
    param(
        [string]$Username,
        [string]$ExportPath
    )
    
    Write-VerboseWithTime "Exportiere Benutzervorlage für: $Username"
    try {
        # Get all user properties
        $user = Get-ADUser -Identity $Username -Properties *
        
        # Get group memberships
        $groups = Get-UserGroupMemberships -Username $Username
        
        # Create custom object with relevant properties
        $exportObject = [PSCustomObject]@{
            SamAccountName = $user.SamAccountName
            UserPrincipalName = $user.UserPrincipalName
            GivenName = $user.GivenName
            Surname = $user.Surname
            Name = $user.Name
            DisplayName = $user.DisplayName
            EmailAddress = $user.EmailAddress
            Department = $user.Department
            Title = $user.Title
            Description = $user.Description
            Groups = $groups -join ';'
            OU = $user.DistinguishedName -replace '^[^,]*,',''
        }
        
        # Set default export path if not provided
        if (-not $ExportPath) {
            $ExportPath = ".\$($Username)_template.csv"
        }
        
        # Export to CSV with German locale (semicolon separator)
        $exportObject | Export-Csv -Path $ExportPath -NoTypeInformation -Delimiter ';' -Encoding UTF8
        Write-VerboseWithTime "Template exported to: $ExportPath"
    }
    catch {
        Write-Error "Failed to export template: $_"
    }
}

# Function to process CSV file
function Process-CSVFile {
    param(
        [string]$CSVPath,
        [string]$TemplateUser,
        [SecureString]$Password,
        [string]$TargetOU
    )
    
    Write-VerboseWithTime "Verarbeite CSV-Datei: $CSVPath"
    try {
        # Import CSV with German locale (semicolon separator)
        $users = Import-Csv -Path $CSVPath -Delimiter ';' -Encoding UTF8
        
        foreach ($user in $users) {
            $userProps = @{
                SamAccountName = $user.SamAccountName
                UserPrincipalName = $user.UserPrincipalName
                GivenName = $user.GivenName
                Surname = $user.Surname
                Name = $user.Name
                DisplayName = $user.DisplayName
                EmailAddress = $user.EmailAddress
                Department = $user.Department
                Title = $user.Title
                Description = $user.Description
            }
            
            # Validate mandatory properties
            $missingProps = Test-MandatoryProperties -Properties $userProps
            if ($missingProps.Count -gt 0) {
                Write-Warning "Skipping user $($user.SamAccountName): Missing mandatory properties: $($missingProps -join ', ')"
                continue
            }
            
            New-UserFromTemplate -TemplateUser $TemplateUser -NewUserProperties $userProps -Password $Password -TargetOU $TargetOU
        }
    }
    catch {
        Write-Error "Failed to process CSV file: $_"
    }
}

# Function to create new user from template
function New-UserFromTemplate {
    param(
        [string]$TemplateUser,
        [hashtable]$NewUserProperties,
        [SecureString]$Password,
        [string]$TargetOU
    )
    
    Write-VerboseWithTime "Erstelle neuen Benutzer basierend auf Vorlage: $TemplateUser"
    try {
        # Get template user properties
        $templateUserObj = Get-ADUser -Identity $TemplateUser -Properties *
        
        # Get group memberships
        $groups = Get-UserGroupMemberships -Username $TemplateUser
        
        # Create new user
        $newUserParams = @{
            Instance = $templateUserObj
            Path = $TargetOU
            AccountPassword = $Password
            Enabled = $true
        }
        
        # Add all new user properties
        foreach ($prop in $NewUserProperties.GetEnumerator()) {
            $newUserParams[$prop.Key] = $prop.Value
        }
        
        # Create the user
        New-ADUser @newUserParams
        
        # Add group memberships
        foreach ($group in $groups) {
            Add-ADGroupMember -Identity $group -Members $NewUserProperties.SamAccountName
        }
        
        Write-VerboseWithTime "Benutzer $($NewUserProperties.SamAccountName) erfolgreich erstellt"
    }
    catch {
        Write-Error "Failed to create user: $_"
    }
}

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
    
    # Get password if not provided
    if (-not $Password) {
        $Password = Get-SecurePassword
    }
    
    # Get target OU if not provided
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
