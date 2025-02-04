# Advanced AD User Management Script
# Author: Bolt
# Version: 2.0
# Description: Creates new AD users based on template users with enhanced error handling and flexibility
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
    Write-Verbose "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Message"
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
            Write-Warning "This field is required. Please enter a value."
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
    
    Write-VerboseWithTime "Exporting template user '$Username' to CSV"
    
    try {
        $user = Get-ADUser -Identity $Username -Properties *
        if (-not $user) {
            throw "Template user not found"
        }
        
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
            'EmailAddress'
        )
        
        $userProps = [PSCustomObject]@{}
        foreach ($prop in $exportProperties) {
            $userProps | Add-Member -MemberType NoteProperty -Name $prop -Value $user.$prop
        }
        
        $fileName = if ([string]::IsNullOrWhiteSpace($ExportPath)) {
            "UserTemplate_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        } else {
            $ExportPath
        }
        
        $userProps | Export-Csv -Path $fileName -NoTypeInformation -Encoding UTF8
        Write-VerboseWithTime "Template exported successfully to: $fileName"
    }
    catch {
        Write-Error "Failed to export template: $_"
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
    
    Write-VerboseWithTime "Creating new user based on template: $TemplateUser"
    
    try {
        # Get template user
        $template = Get-ADUser -Identity $TemplateUser -Properties *
        if (-not $template) {
            throw "Template user not found"
        }
        
        # Create new user properties
        $userProps = @{
            Instance = $template
            Enabled = $true
            ChangePasswordAtLogon = $true
        }
        
        # Add all new user properties
        foreach ($key in $NewUserProperties.Keys) {
            $userProps[$key] = $NewUserProperties[$key]
        }
        
        # Create the new user
        New-ADUser @userProps
        
        Write-VerboseWithTime "Successfully created user: $($NewUserProperties.SamAccountName)"
        return $true
    }
    catch {
        Write-Error "Failed to create user: $_"
        return $false
    }
}

# Function to process CSV file
function Process-CSVFile {
    param(
        [string]$CSVPath,
        [string]$TemplateUser
    )
    
    Write-VerboseWithTime "Processing CSV file: $CSVPath"
    
    try {
        $users = Import-Csv -Path $CSVPath
        foreach ($user in $users) {
            $userProps = @{}
            foreach ($prop in $user.PSObject.Properties) {
                if (-not [string]::IsNullOrWhiteSpace($prop.Value)) {
                    $userProps[$prop.Name] = $prop.Value
                }
            }
            
            $missingProps = Test-MandatoryProperties -Properties $userProps
            if ($missingProps.Count -gt 0) {
                Write-Warning "Missing mandatory properties for user: $($missingProps -join ', ')"
                continue
            }
            
            New-UserFromTemplate -TemplateUser $TemplateUser -NewUserProperties $userProps
        }
    }
    catch {
        Write-Error "Failed to process CSV file: $_"
        return $false
    }
    
    return $true
}

# Main script logic
try {
    # Set verbose output
    $VerbosePreference = "Continue"
    
    Write-VerboseWithTime "Starting AD User Management Script"
    
    # Handle template export
    if ($ExportTemplate) {
        if ([string]::IsNullOrWhiteSpace($TemplateUser)) {
            $TemplateUser = Get-ValidatedInput -Prompt "Enter template username" -Required
        }
        Export-UserTemplate -Username $TemplateUser -ExportPath $ExportPath
        return
    }
    
    # Interactive mode if no parameters provided
    if (-not $TemplateUser -and -not $CSV) {
        Write-VerboseWithTime "Starting interactive mode"
        $TemplateUser = Get-ValidatedInput -Prompt "Enter template username" -Required
        
        $mode = Get-ValidatedInput -Prompt "Enter mode (single/csv)" -Required
        if ($mode -eq "csv") {
            $CSV = Get-ValidatedInput -Prompt "Enter CSV path" -Required
        } else {
            $userProps = @{
                SamAccountName = Get-ValidatedInput -Prompt "Enter new username (SAMAccountName)" -Required
                UserPrincipalName = Get-ValidatedInput -Prompt "Enter UserPrincipalName" -Required
                GivenName = Get-ValidatedInput -Prompt "Enter first name" -Required
                Surname = Get-ValidatedInput -Prompt "Enter last name" -Required
                DisplayName = Get-ValidatedInput -Prompt "Enter display name"
                Description = Get-ValidatedInput -Prompt "Enter description"
            }
            
            New-UserFromTemplate -TemplateUser $TemplateUser -NewUserProperties $userProps
        }
    }
    # CSV mode
    elseif ($CSV) {
        if (-not $TemplateUser) {
            $TemplateUser = Get-ValidatedInput -Prompt "Enter template username" -Required
        }
        Process-CSVFile -CSVPath $CSV -TemplateUser $TemplateUser
    }
    # Parameter mode
    else {
        if (-not $NewUserName) {
            $NewUserName = Get-ValidatedInput -Prompt "Enter new username (SAMAccountName)" -Required
        }
        
        $userProps = @{
            SamAccountName = $NewUserName
            UserPrincipalName = Get-ValidatedInput -Prompt "Enter UserPrincipalName" -Required
            GivenName = Get-ValidatedInput -Prompt "Enter first name" -Required
            Surname = Get-ValidatedInput -Prompt "Enter last name" -Required
            DisplayName = Get-ValidatedInput -Prompt "Enter display name"
            Description = Get-ValidatedInput -Prompt "Enter description"
        }
        
        New-UserFromTemplate -TemplateUser $TemplateUser -NewUserProperties $userProps
    }
    
    Write-VerboseWithTime "Script completed successfully"
}
catch {
    Write-Error "Script failed: $_"
    Write-VerboseWithTime "Script terminated with errors"
}
finally {
    $VerbosePreference = "SilentlyContinue"
}
