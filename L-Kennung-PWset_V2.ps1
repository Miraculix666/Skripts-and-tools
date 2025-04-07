# Import Active Directory Module
Import-Module ActiveDirectory

# Debug-Funktion für OU-Suche
function Find-TargetOUs {
    param($Names)
    
    Write-Host "`n[DEBUG] Starte OU-Suche nach Namen: $Names"
    $domain = Get-ADDomain
    Write-Host "Domain-DN: $($domain.DistinguishedName)"
    
    $foundOUs = @()
    foreach ($name in $Names) {
        $ous = Get-ADOrganizationalUnit -Filter "Name -like '*$name*'" -SearchBase $domain.DistinguishedName -SearchScope Subtree
        if ($ous) {
            Write-Host "Gefundene OUs für '$name':"
            $ous | Format-Table Name, DistinguishedName -AutoSize
            $foundOUs += $ous
        }
        else {
            Write-Host "Warnung: Keine OUs gefunden für '$name'" -ForegroundColor Yellow
        }
    }
    
    return $foundOUs | Sort-Object DistinguishedName -Unique
}

# OU-Suche mit Debug-Informationen
$targetOUs = Find-TargetOUs -Names @('81', '82')

if (-not $targetOUs) {
    Write-Host "Fehler: Ziel-OUs konnten nicht gefunden werden!" -ForegroundColor Red
    exit
}

# Rest des Skripts wie zuvor...
