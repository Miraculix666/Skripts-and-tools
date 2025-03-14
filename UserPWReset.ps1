# Define parameters
param (
    [string]$UserName,
    [string]$Wildcard,
    [string]$OU,
    [string]$Password,
    [switch]$SkipConfirmation = $false,
    [switch]$VerboseOutput = $true
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
        
        Write-Host "Password reset and account enabled for $Identity" -ForegroundColor Green
    } catch {
        Write-Host "Failed to reset password and enable account for $Identity" -ForegroundColor Red
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

# Find users based on input
if ($UserName) {
    $users = Get-ADUser -Identity $UserName
} elseif ($Wildcard) {
    $users = Get-ADUser -Filter {Name -like $Wildcard}
} elseif ($OU) {
    $users = Get-ADUser -Filter * -SearchBase $OU
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

# Process each user
foreach ($user in $users) {
    ResetPasswordAndEnableAccount -Identity $user.SamAccountName -NewPassword $Password
}
