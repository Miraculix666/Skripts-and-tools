# Define the distinguished names of the Organizational Units (OUs)
$ou81 = "OU=81,OU=Polizei-NRW-PB-PE-2012,DC=polizei,DC=nrw,DC=de"
$ou82 = "OU=82,OU=Polizei-NRW-PB-PE-2012,DC=polizei,DC=nrw,DC=de"

# Define the output file paths
$outputFilePath = "C:\Daten\Users_LastLogon_Report.txt"
$samAccountNamesFilePath = "C:\Daten\All_SAMAccountNames.txt"
$expiredUsersFilePath = "C:\Daten\Expired_Users_SAM.txt"
$usersWithoutEmailFilePath = "C:\Daten\Users_Without_Email.txt"

# Get users from OU 81 with "L110" or "L114" in the username
$usersOU81 = Get-ADUser -Filter {Enabled -eq $true -and (Name -like "L110*" -or Name -like "L114*")} -SearchBase $ou81

# Get users from OU 82 with "L110" or "L114" in the username
$usersOU82 = Get-ADUser -Filter {Enabled -eq $true -and (Name -like "L110*" -or Name -like "L114*")} -SearchBase $ou82

# Combine the results from both OUs
$allUsers = $usersOU81 + $usersOU82

# Display the results
if ($allUsers.Count -gt 0
