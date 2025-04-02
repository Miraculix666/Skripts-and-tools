# Active Directory-Modul importieren
Import-Module ActiveDirectory

# OU ermitteln
$ou = Get-ADOrganizationalUnit -Identity 'OU=81,DC=IhrDomain,DC=com'

# Distinguished Name (DN) der OU
$dn = $ou.DistinguishedName

# Partition ermitteln (standardmäßig die Domain-Partition)
$partition = "DC=" + $dn.Split(",")[1].Split("=")[1] + "," + $dn.Split(",")[2].Split("=")[1]

# Kontext ermitteln (aus dem DN)
$context = $dn

Write-Host "Distinguished Name (DN): $dn"
Write-Host "Partition: $partition"
Write-Host "Kontext: $context"

# Ablaufdaten der Passwörter abrufen und exportieren
Get-ADUser -Filter * -SearchBase $dn -Properties "DisplayName", "msDS-UserPasswordExpiryTimeComputed" |
Select-Object -Property "DisplayName", @{Name="ExpiryDate";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}} |
Export-Csv -Path "C:\temp\OU81_PasswordExpiration.csv" -NoTypeInformation
