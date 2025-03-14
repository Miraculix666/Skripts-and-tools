# Define parameters
param (
    [string]$UserName,
    [string]$Wildcard,
    [string]$OU,
    [string]$Password,
    [switch]$SkipConfirmation = $false,
    [switch]$VerboseOutput = $true,
    [switch]$Silent = $false
)

# Load necessary modules
Import-Module ActiveDirectory

# Function to reset password and enable account
function ResetPasswordAndEnableAccount {
    param (
        [string]$Identity,
        [string]$NewPassword
    )
    
    try {
        # Convert password to secure string
        $securePassword = ConvertTo-SecureString -String $NewPassword -AsPlainText -Force
        
        # Reset password and enable account
        Set-ADAccountPassword -Identity $Identity -Reset -NewPassword $securePassword -ErrorAction Stop
        Enable-ADAccount -Identity $Identity -ErrorAction Stop
        Set-ADAccountControl -Identity $Identity -CannotChangePassword $true -ErrorAction Stop
        
        if (!$Silent) {
            Write-Host "Password reset and account enabled for $Identity" -ForegroundColor Green
        }
        Write-Verbose "Password reset and account enabled for $Identity"
    } catch {
        if (!$Silent) {
            Write-Host "Failed to reset password and enable account for $Identity" -ForegroundColor Red
        }
        Write-Verbose "Failed to reset password and enable account for $Identity"
    }
}

# Main script logic
if (!$UserName -and !$Wildcard -and !$OU) {
    # Interactive mode
    $UserName = Read-Host "Enter user name (leave blank for wildcard or OU search)"
    $Wildcard = Read-Host "Enter wildcard (e.g., DWR45*)"
    $OU = Read-Host "Enter OU (e.g., OU=47678)"
    $Password = Read-Host "Enter new password"
} elseif (!$Password) {
    $Password = Read-Host "Enter new password"
}

# Determine domain name automatically
$domainName = (Get-CimInstance Win32_ComputerSystem).Domain
Write-Verbose "Detected domain name: $domainName"

# Find users based on input
$users = @()
if ($UserName) {
    $users += Get-ADUser -Identity $UserName
} elseif ($Wildcard) {
    # Use the wildcard in the filter
    $allUsers = Get-ADUser -Filter * -Properties Name
    $users = $allUsers | Where-Object {$_.Name -like $Wildcard}
} elseif ($OU) {
    # Get all users in the specified OU
    $users += Get-ADUser -Filter * -SearchBase $OU
}

# Prompt for confirmation before changing
if (!$SkipConfirmation) {
    Write-Host "The following users will be affected:"
    $users | Select-Object Name, SamAccountName, DistinguishedName | Format-Table -AutoSize
    $confirm = Read-Host "Are these the correct users? (yes/no)"
    if ($confirm -ne "yes") {
        Write-Host "Operation cancelled."
        exit
    }
}

# Save users to file with German locale settings
$users | Select-Object Name, SamAccountName, DistinguishedName | Export-Csv -Path "C:\Daten\AffectedUsers.csv" -NoTypeInformation -Delimiter ";"

# Process each user
foreach ($user in $users) {
    ResetPasswordAndEnableAccount -Identity $user.SamAccountName -NewPassword $Password
}

# Logging
Start-Transcript -Path "C:\Daten\ScriptLog.log"
Write-Host "Script execution completed."
Stop-Transcript

# Help function
function Get-Help {
    Write-Host "Usage: .\Script.ps1 [-UserName <string>] [-Wildcard <string>] [-OU <string>] [-Password <string>] [-SkipConfirmation] [-VerboseOutput] [-Silent]"
    Write-Host "Parameters:"
    Write-Host "  -UserName      : Specify a user name"
    Write-Host "  -Wildcard      : Specify a wildcard for user names"
    Write-Host "  -OU            : Specify an Organizational Unit"
    Write-Host "  -Password      : Specify a new password"
    Write-Host "  -SkipConfirmation : Skip user confirmation"
    Write-Host "  -VerboseOutput : Enable verbose output (default)"
    Write-Host "  -Silent        : Suppress console output"
}

# Call help function if no parameters are provided
if (!$UserName -and !$Wildcard -and !$OU -and !$Password) {
    Get-Help
}
