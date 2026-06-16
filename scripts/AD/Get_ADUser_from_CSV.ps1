param (
    [string]$InputCsvPath,
    [string]$OutputCsvPath
)

# Importieren der CSV-Informationen ohne Kopfzeile
$CSVImport = Get-Content $InputCsvPath | ConvertFrom-Csv -Delimiter "," -Header "Nachname", "Vorname"

# Für jeden Datensatz im CSV
foreach ($Benutzer in $CSVImport) {
    $Nachname = $Benutzer.Nachname
    $Vorname = $Benutzer.Vorname
    $SamAccountName = Get-ADUser -Filter { Surname -eq $Nachname -and GivenName -eq $Vorname } | Select-Object -ExpandProperty SamAccountName
    [PSCustomObject]@{
        Nachname = $Nachname
        Vorname = $Vorname
        SamAccountName = $SamAccountName
    } | Export-Csv -Append -Path $OutputCsvPath -NoTypeInformation
}

# Erfolgsmeldung
Write-Host "Die SamAccountNamen wurden erfolgreich in $OutputCsvPath gespeichert."
