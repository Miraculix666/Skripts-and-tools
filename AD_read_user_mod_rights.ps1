param(
    [string]$OU_DN = "OU=BeispielOU,DC=deineDomäne,DC=de",
    [string]$Identity = "DOMAIN\\Benutzername",
    [string]$TestUser = "<SamAccountName_eines_Benutzers_in_der_OU>"
)

# Active Directory Modul laden
Import-Module ActiveDirectory

# Alle Felder des Benutzerobjekts dynamisch bestimmen
$allAttributes = (Get-ADUser -Identity $TestUser -Properties * | Get-Member -MemberType Property).Name

# ACL der OU auslesen
$acl = Get-Acl -Path ("AD:\" + $OU_DN)
$effectiveRights = $acl.Access | Where-Object {
    $_.IdentityReference -eq $Identity -and $_.ActiveDirectoryRights -match "WriteProperty|CreateChild"
}

# Ergebnis-Tabelle
$result = @()

foreach ($attr in $allAttributes) {
    $hasRight = $effectiveRights | Where-Object { $_.ObjectType -eq $attr }
    $obj = [PSCustomObject]@{
        Attribut = $attr
        Schreibrecht = if ($hasRight) {"Ja"} else {"Nein"}
    }
    $result += $obj
}

# Ausgabe als Tabelle
$result | Format-Table -AutoSize
