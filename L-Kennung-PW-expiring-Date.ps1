

# Active Directory-Modul importieren
Import-Module ActiveDirectory

# Aktuelle Domäne ermitteln
$aktuelleDomäne = Get-ADDomain

Write-Host "Aktuelle Domäne: $($aktuelleDomäne.DNSRoot)"

# Definieren der OUs
$ou81 = "OU=81,DC=IhrDomain,DC=com"
$ou82 = "OU=82,DC=IhrDomain,DC=com"

# Ablaufdaten der Passwörter abrufen und exportieren
Get-ADUser -Filter * -SearchBase $ou81 -Properties "DisplayName", "msDS-UserPasswordExpiryTimeComputed" |
Select-Object -Property "DisplayName", @{Name="ExpiryDate";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}} |
Export-Csv -Path "C:\temp\OU81_PasswordExpiration.csv" -NoTypeInformation

Get-ADUser -Filter * -SearchBase $ou82 -Properties "DisplayName", "msDS-UserPasswordExpiryTimeComputed" |
Select-Object -Property "DisplayName", @{Name="ExpiryDate";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}} |
Export-Csv -Path "C:\temp\OU82_PasswordExpiration.csv" -NoTypeInformation
