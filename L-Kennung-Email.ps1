# Define the distinguished names of the Organizational Units (OUs)
$ou81 = "OU=81,OU=Polizei-NRW-PB-PE-2012,DC=polizei,DC=nrw,DC=de"
$ou82 = "OU=82,OU=Polizei-NRW-PB-PE-2012,DC=polizei,DC=nrw,DC=de"

# Define the output file paths
$outputFilePath = "C:\Daten\Users_LastLogon_Report.txt"
$samAccountNamesFilePath = "C:\Daten\All_SAMAccountNames.txt"
$expiredUsersFilePath = "C:\Daten\Expired_Users_SAM.txt"
$emailInfoFilePath = "C:\Daten\Email_Info.txt"
$usersWithoutEmailFilePath = "C:\Daten\Users_Without_Email.txt"

# Get users from OU 81 with "L110" or "L114" in the username
$usersOU81 = Get-ADUser -Filter {Enabled -eq $true -and (Name -like "L110*" -or Name -like "L114*")} -SearchBase $ou81

# Get users from OU 82 with "L110" or "L114" in the username
$usersOU82 = Get-ADUser -Filter {Enabled -eq $true -and (Name -like "L110*" -or Name -like "L114*")} -SearchBase $ou82

# Combine the results from both OUs
$allUsers = $usersOU81 + $usersOU82

# Display the results
if ($allUsers.Count -gt 0) {
    Write-Host "Users in specified OUs with 'L110' or 'L114' in the username:"
    $allUsers | Select-Object Name, SamAccountName, DistinguishedName | Format-Table -AutoSize

    # Write all SAMAccountNames in the given scope into a file
    $allUsers | ForEach-Object { $_.SamAccountName | Out-File -Append -FilePath $samAccountNamesFilePath }

    # Get all users and their last logon date
    $usersLastLogon = $allUsers | Get-ADUser -Properties LastLogonDate, EmailAddress | Select-Object Name, SamAccountName, LastLogonDate, EmailAddress

    # Sort users by the oldest last logon date first
    $usersLastLogon = $usersLastLogon | Sort-Object LastLogonDate

    # Calculate the date 39 weeks ago
    $thresholdDate = (Get-Date).AddDays(-39 * 7)

    # Format the results and mark dates older than 39 weeks 
    $results = $usersLastLogon | ForEach-Object {
        $lastLogonDate = $_.LastLogonDate
        if ($lastLogonDate -ne $null) {
            $formattedDate = Get-Date $lastLogonDate -Format "yyyy-MM-dd HH:mm:ss"
            $formattedDatemarked = if ($lastLogonDate -lt $thresholdDate) {  "xxxxx $formattedDate" } else { $formattedDate }
        } else {
            $formattedDatemarked = "xxxxx N/A"
        }
        [PSCustomObject]@{
            Name            = $_.Name
            SamAccountName  = $_.SamAccountName
            LastLogonDate   = $formattedDatemarked
            EmailAddress    = $_.EmailAddress
        }
    }

    # Write all SAMAccountNames with logon older than 39 weeks into a file
    $expiredUsers = $results | Where-Object { $_.LastLogonDate -like "xxxxx*" } | ForEach-Object { $_.SamAccountName | Out-File -Append -FilePath $expiredUsersFilePath }

    # Display the results
    if ($results.Count -gt 0) {
        Write-Host "Users and their last logon dates:"
        $results | Format-Table -AutoSize

        # Write results to a file
        $results | Export-Csv -Path $outputFilePath -NoTypeInformation
        Write-Host "Output saved to: $outputFilePath"

        # Separate users without email addresses
        $usersWithoutEmail = $results | Where-Object { $_.EmailAddress -eq $null }
        $usersWithoutEmail | Select-Object Name, SamAccountName | Format-Table -AutoSize | Out-File -FilePath $usersWithoutEmailFilePath
        Write-Host "Users without email addresses written to: $usersWithoutEmailFilePath"

        # Write SAMAccountNames and associated email addresses to a file
        $results | Where-Object { $_.EmailAddress -ne $null } | Select-Object SamAccountName, EmailAddress | Format-Table -AutoSize | Out-File -FilePath $emailInfoFilePath
        Write-Host "SAMAccountNames and associated email addresses written to: $emailInfoFilePath"
    } else {
        Write-Host "No users found with last logon dates."
    }
} else {
    Write-Host "No users found in the specified OUs with 'L110' or 'L114' in the username."
}

Write-Host "All SAMAccountNames have been written to: $samAccountNamesFilePath"
Write-Host "Expired SAMAccountNames have been written to: $expiredUsersFilePath"
