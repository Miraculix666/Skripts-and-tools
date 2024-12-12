
Get-ADUser -Filter "Name -like 'L*'" -Properties Name, SamAccountName, MemberOf | Export-Csv -Path "C:\AD_User_L.csv" -NoTypeInformation
