# Pfad zur Eingabe-CSV-Datei (Nachname,Vorname)
$InputCsvPath = "C:\Pfad\Zur\Deiner\Eingabe-Datei.csv"

# Pfad zur Ausgabe-CSV-Datei (Ergebnisse speichern)
$OutputCsvPath = "C:\Pfad\Zur\Deiner\Ausgabe-Datei.csv"

# Importieren der CSV-Informationen
$CSVImport = Import-Csv $InputCsvPath -Delimiter "," -Encoding Default

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
