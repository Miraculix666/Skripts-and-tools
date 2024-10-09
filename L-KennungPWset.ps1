Gesendet von Copilot:
Natürlich! Hier ist das angepasste Skript, das eine Variable für das Passwort verwendet:

# Define the distinguished names of the Organizational Units (OUs)
$ou81 = "OU=81,OU=Polizei-NRW-PB-PE-2012,DC=polizei,DC=nrw,DC=de"
$ou82 = "OU=82,OU=Polizei-NRW-PB-PE-2012,DC=polizei,DC=nrw,DC=de"

# Define the output file paths
$outputFilePath = "C:\Daten\Users_LastLogon_Report.txt"
$samAccountNamesFilePath = "C:\Daten\All_SAMAccountNames.txt"
$expiredUsersFilePath = "C:\Daten\Expired_Users_SAM.txt"

# Get users from OU 81 with "L110" or "L114" in the username
$usersOU81 = Get-ADUser -Filter {Enabled -eq $true -and (Name -like "L110*" -or Name -like "L114*")} -SearchBase $ou81

# Get users from OU 82 with "L110" or "L114" in the username
$usersOU82 = Get-ADUser -Filter {Enabled -eq $true -and (Name -like "L110*" -or Name -like "L114*")} -SearchBase $ou82

# Combine the results from both OUs
$allUsers = $usersOU81 + $usersOU82

# Input file path containing the list of affected users
$inputFilePath = "C:\Daten\Deaktivierte_L_Kennung_SAM.txt"

# Define the new password
$passwordString = "P2f7aL4!01"
$newPassword = ConvertTo-SecureString -String $passwordString -AsPlainText -Force

# Read SamAccountNames from the file
$samAccountNames = Get-Content -Path $inputFilePath

# Reset password for each user
foreach ($samAccountName in $samAccountNames) {
    try {
        Set-ADAccountPassword -Identity $samAccountName -Reset -NewPassword $newPassword -ErrorAction Stop
        Write-Host "Password reset successful for user: $samAccountName" -ForegroundColor Green
    } catch {
        Write-Host "Failed to reset password for user: $samAccountName" -ForegroundColor Red
    }
}

Write-Host "Password reset completed for users in the file."

Jetzt wird das Passwort in einer Variable gespeichert und verwendet. Wenn du weitere Anpassungen benötigst, lass es mich wissen!
