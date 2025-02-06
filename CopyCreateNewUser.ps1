# Copy-ADUserTemplate.ps1
#
# .SYNOPSIS
# Creates new Active Directory users based on a template user account.
#
# .DESCRIPTION
# This script clones an existing Active Directory user as a template for creating new users.
# It can operate in interactive mode or batch mode using a CSV file.
# All AD properties, group memberships, and privileges are copied from the template user.
#
# .PARAMETER TemplateUserDN
# Distinguished Name of the template user to copy from
#
# .PARAMETER CsvPath
# Optional path to CSV file containing new user information
#
# .PARAMETER TargetOU
# Optional Distinguished Name of target OU for new users
#
# .PARAMETER LogPath
# Path for log file. Defaults to ".\ADUserClone_<timestamp>.log"
#
# .EXAMPLE
# .\Copy-ADUserTemplate.ps1 -TemplateUserDN "CN=Template User,OU=Users,DC=contoso,DC=com"
# Creates new users interactively based on the specified template
#
# .EXAMPLE
# .\Copy-ADUserTemplate.ps1 -TemplateUserDN "CN=Template User,OU=Users,DC=contoso,DC=com" -CsvPath ".\new_users.csv"
# Creates multiple users from CSV file using the specified template
#
# .NOTES
# Version: 1.0
# Author: Bolt
# Requires: Active Directory PowerShell Module
# Compatible with PowerShell 5.1
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$TemplateUserDN,
    
    [Parameter(Mandatory = $false)]
    [string]$CsvPath,
    
    [Parameter(Mandatory = $false)]
    [string]$TargetOU,
    
    [Parameter(Mandatory = $false)]
    [string]$LogPath = ".\ADUserClone_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
)

# Import required module
Import-Module ActiveDirectory

# Initialize logging
function Write-Log {
    param($Message)
    $logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Message"
    Write-Verbose $logMessage
    Add-Content -Path $LogPath -Value $logMessage
}

# Mandatory attributes that must be unique for each user
$mandatoryAttributes = @(
    'sAMAccountName',
    'userPrincipalName',
    'givenName',
    'sn'
)

# Optional attributes to copy from template
$optionalAttributes = @(
    'displayName',
    'description',
    'mail',
    'physicalDeliveryOfficeName',
    'department',
    'company',
    'title',
    'streetAddress',
    'l',
    'postalCode',
    'c',
    'telephoneNumber',
    'mobile',
    'homePhone',
    'wWWHomePage',
    'employeeID'
)

function Get-SecurePassword {
    $password = Read-Host "Enter password for new user" -AsSecureString
    return $password
}

function Get-TemplateUser {
    param([string]$DN)
    try {
        Write-Log "Retrieving template user: $DN"
        $user = Get-ADUser -Identity $DN -Properties *
        if (-not $user) {
            throw "Template user not found"
        }
        return $user
    }
    catch {
        Write-Log "Error retrieving template user: $_"
        throw
    }
}

function Get-UserGroups {
    param($TemplateUser)
    try {
        Write-Log "Retrieving group memberships for template user"
        return Get-ADUser $TemplateUser -Properties MemberOf | Select-Object -ExpandProperty MemberOf
    }
    catch {
        Write-Log "Error retrieving group memberships: $_"
        throw
    }
}

function New-ClonedUser {
    param(
        $TemplateUser,
        $NewUserProps,
        $Groups,
        $TargetOU
    )
    
    try {
        Write-Log "Creating new user with sAMAccountName: $($NewUserProps.sAMAccountName)"
        
        # Prepare user properties
        $userParams = @{
            Instance = $TemplateUser
            Path = if ($TargetOU) { $TargetOU } else { ($TemplateUser.DistinguishedName -split ',', 2)[1] }
            Enabled = $true
        }
        
        # Add mandatory properties
        foreach ($attr in $mandatoryAttributes) {
            if ($NewUserProps.$attr) {
                $userParams[$attr] = $NewUserProps.$attr
            }
        }
        
        # Add optional properties
        foreach ($attr in $optionalAttributes) {
            if ($NewUserProps.$attr) {
                $userParams[$attr] = $NewUserProps.$attr
            }
        }
        
        # Create user
        $newUser = New-ADUser @userParams -PassThru
        
        # Set password
        $password = if ($NewUserProps.Password) {
            ConvertTo-SecureString $NewUserProps.Password -AsPlainText -Force
        } else {
            Get-SecurePassword
        }
        Set-ADAccountPassword -Identity $newUser -NewPassword $password
        
        # Add group memberships
        foreach ($group in $Groups) {
            Add-ADGroupMember -Identity $group -Members $newUser
            Write-Log "Added user to group: $group"
        }
        
        Write-Log "Successfully created user: $($NewUserProps.sAMAccountName)"
        return $newUser
    }
    catch {
        Write-Log "Error creating user: $_"
        throw
    }
}

function Process-InteractiveMode {
    param(
        $TemplateUser,
        $Groups
    )
    
    try {
        Write-Log "Starting interactive mode"
        
        $newUserProps = @{}
        
        # Get mandatory properties
        foreach ($attr in $mandatoryAttributes) {
            $newUserProps[$attr] = Read-Host "Enter $attr"
        }
        
        # Get optional properties (with default from template)
        foreach ($attr in $optionalAttributes) {
            $default = $TemplateUser.$attr
            $input = Read-Host "Enter $attr (default: $default)"
            $newUserProps[$attr] = if ($input) { $input } else { $default }
        }
        
        New-ClonedUser -TemplateUser $TemplateUser -NewUserProps $newUserProps -Groups $Groups -TargetOU $TargetOU
    }
    catch {
        Write-Log "Error in interactive mode: $_"
        throw
    }
}

function Process-CsvMode {
    param(
        $TemplateUser,
        $Groups,
        $CsvPath
    )
    
    try {
        Write-Log "Starting CSV mode with file: $CsvPath"
        
        # Import CSV with German locale (semicolon separator)
        $users = Import-Csv -Path $CsvPath -Delimiter ';'
        
        foreach ($user in $users) {
            # Verify mandatory attributes
            foreach ($attr in $mandatoryAttributes) {
                if (-not $user.$attr) {
                    throw "Missing mandatory attribute '$attr' in CSV for user"
                }
            }
            
            New-ClonedUser -TemplateUser $TemplateUser -NewUserProps $user -Groups $Groups -TargetOU $TargetOU
        }
    }
    catch {
        Write-Log "Error in CSV mode: $_"
        throw
    }
}

# Main script execution
try {
    Write-Log "Script started"
    
    # Get template user
    $templateUser = Get-TemplateUser -DN $TemplateUserDN
    Write-Log "Template user found: $($templateUser.sAMAccountName)"
    
    # Get group memberships
    $groups = Get-UserGroups -TemplateUser $templateUser
    Write-Log "Retrieved ${groups.Count} group memberships"
    
    # Process users
    if ($CsvPath) {
        Process-CsvMode -TemplateUser $templateUser -Groups $groups -CsvPath $CsvPath
    }
    else {
        Process-InteractiveMode -TemplateUser $templateUser -Groups $groups
    }
    
    Write-Log "Script completed successfully"
}
catch {
    Write-Log "Fatal error: $_"
    throw
}
