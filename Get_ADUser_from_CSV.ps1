# Pfad zur CSV-Datei (kann als Argument übergeben werden)
param (
    [string]$Import-CSV
)

# Importieren der CSV-Informationen
$CSVImport = Import-Csv $CsvPath -Delimiter ";" -Encoding Default

# Für jeden Datensatz im CSV
foreach ($Benutzer in $CSVImport) {
    # Active Directory Benutzer erstellen
    New-ADUser -Path "OU=NeueBenutzer,OU=Benutzer,OU=GPS,DC=gps,DC=germanpowershell,DC=com" `
        -Surname $Benutzer.Name `
        -GivenName $Benutzer.Vorname `
        -SamAccountName $Benutzer.Login `
        -UserPrincipalName $Benutzer.Login `
        -AccountPassword ($Benutzer.Passwort | ConvertTo-SecureString -AsPlainText -Force) `
        -Enabled:$true `
        -DisplayName "$($Benutzer.Vorname) $($Benutzer.Name)" `
        -Name "$($Benutzer.Vorname) $($Benutzer.Name)"
}
