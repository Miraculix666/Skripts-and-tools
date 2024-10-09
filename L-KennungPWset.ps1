
#####ACHUTNG, funzt nicht!! erwischt falsche kennungen.
##### neuer Ansatztz mit Filter OU 81 & 82

# Set the distinguished names of the Organizational Units (OU)
$ou81 = "OU=Benutzer,OU=81,OU=Polizei-NRW-PB-PE-2012,DC=polizei,DC=nrw,dc=de"
$ou82 = "OU=Benutzer,OU=82,OU=Polizei-NRW-PB-PE-2012,DC=polizei,DC=nrw,dc=de"

# Define the output file path
$outputFilePath = "c:\Daten\L_Kennung.csv"

# Get all deactivated users in OU 81 with "L110" or "L114" in the username
$UsersOU81 = Get-ADUser -Filter {Name -like "L110*" -or Name -like "L114*"} -SearchBase $ou81

# Get all deactivated users in OU 82 with  "L110" or "L114" in the username
$UsersOU82 = Get-ADUser -Filter {Name -like "L110*" -or Name -like "L114*"} -SearchBase $ou82

# Combine the results from both OUs
$allUsers = $UsersOU81 + $UsersOU82

# Display the results and write to CSV file
if ($allUsers.Count -gt 0) {
    Write-Host "Users in specified search path with names containing '110' or '114':"
    $allUsers | Select-Object Name, SamAccountName, DistinguishedName
    $allUsers | Select-Object Name, SamAccountName, DistinguishedName | Export-Csv -Path $outputFilePath -NoTypeInformation
    Write-Host "Output saved to: $outputFilePath"}

# Write SamAccountNames to the file
$allUsers | ForEach-Object { $_.SamAccountName | Out-File -Append -FilePath $outputFilePathSAM }
# Display message about the file
Write-Host "SamAccountNames have been written to: $outputFilePathSAM"

# Input file path containing the list of affected users
$inputFilePath = "c:\Daten\L_Kennung.csv"

# Set the new password
$newPassword = ConvertTo-SecureString -String "P2f7aL4!10" -AsPlainText -Force

# Read SamAccountNames from the file
$samAccountNames = Get-Content -Path $inputFilePath

# Reset password for each user
foreach ($samAccountName in $samAccountNames) {
    try {
        Set-ADAccountPassword -Identity $samAccountName -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $newPassword -Force) -ErrorAction Stop
        Write-Host "Password reset successful for user: $samAccountName" -ForegroundColor Green
    } catch {
        Write-Host "Failed to reset password for user: $samAccountName" -ForegroundColor Red
    }
}

Write-Host "Password reset completed for users in the file."
