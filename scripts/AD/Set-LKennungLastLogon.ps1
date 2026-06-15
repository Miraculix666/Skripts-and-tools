# Define the path to the file containing SAMAccountNames of expired users
$expiredUsersFilePath = "C:\Daten\Expired_Users.txt"

# Read SAMAccountNames from the file
$expiredUsers = Get-Content -Path $expiredUsersFilePath

# Set the new LastLogonDate
$newLastLogonDate = Get-Date

# Auto-login and set the LastLogonDate for each expired user
foreach ($expiredUser in $expiredUsers) {
    Write-Host "Auto-logging into user: $expiredUser"

    # Assuming $expiredUser contains the SamAccountName, use the following line
    $user = Get-ADUser -Filter { SamAccountName -eq $expiredUser }

    if ($user) {
        # Set the new LastLogonDate for the user
        $user | Set-ADUser -Replace @{LastLogonDate = $newLastLogonDate}

        Write-Host "LastLogonDate updated for user: $expiredUser"
    } else {
        Write-Host "User $expiredUser not found in Active Directory."
