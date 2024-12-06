# Benutzer mit Anmeldenamen beginnend mit "L" abrufen
Write-Verbose "Suche nach AD-Benutzern mit Anmeldenamen beginnend mit 'L'..."
$benutzer = Get-ADUser -Filter "SamAccountName -like 'L*'" -Properties SamAccountName, Name, MemberOf, DistinguishedName

# Ausgabe der gefundenen Benutzer in der Shell
Write-Host "Gefundene Benutzer:"
$benutzer | Format-Table SamAccountName, Name

# Benutzerinformationen, OU und Gruppenzugehörigkeiten in einer Tabelle zusammenfassen
Write-Verbose "Erstelle Tabelle mit Benutzerinformationen, OU und Gruppenzugehörigkeiten..."
$tabelle = @()
$benutzer | ForEach-Object {
    Write-Verbose "Verarbeite Benutzer '$($_.SamAccountName)'..."
    $ou = $_.DistinguishedName -replace '^CN=[^,]+,OU=([^,]+),.*$', '$1'
    # OU "Benutzer" ignorieren
    if ($ou -ne "Benutzer") {
        $_.MemberOf | Get-ADGroup | ForEach-Object {
            $tabelle += [PSCustomObject]@{
                OU = $ou
                Benutzer = $_.Name
                Gruppe = $_.Name # Hier korrekt: $_.Name für Benutzer und Gruppe
            }
            # Ausgabe der aktuellen Benutzer-Gruppen-Zuordnung in der Shell
            Write-Host "Benutzer: $($_.Name), OU: $ou, Gruppe: $($_.Name)"
        }
    }
}

# Tabelle nach OU, Benutzer und dann nach Gruppe sortieren
Write-Verbose "Sortiere Tabelle nach OU, Benutzer und Gruppe..."
$tabelle = $tabelle | Sort-Object -Property OU, Benutzer, Gruppe

# Ausgabe der sortierten Tabelle in der Shell
Write-Host "Sortierte Tabelle:"
$tabelle | Format-Table

# Ergebnis als CSV-Datei speichern
$csvPfad = "C:\daten\AD_Benutzer_Gruppen_L.csv"

# --- Datei löschen, falls vorhanden ---
if (Test-Path $csvPfad) {
    $antwort = Read-Host -Prompt "Die Datei '$csvPfad' existiert bereits. Soll sie gelöscht werden? (Ja/Nein)"
    if ($antwort -eq "Ja") {
        Remove-Item $csvPfad -Force
        Write-Verbose "Datei '$csvPfad' gelöscht."
    }
}

Write-Verbose "Speichere Ergebnis in CSV-Datei '$csvPfad'..."
$tabelle | Export-Csv -Path $csvPfad -NoTypeInformation -Encoding UTF8

# --- Excel-Datei erstellen und formatieren ---

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Worksheets.Item(1)

# CSV-Daten in Excel importieren
$excel.DisplayAlerts = $false 
$range = $worksheet.Range("A1").CurrentRegion
$csv = Get-Content $csvPfad | ConvertFrom-Csv
$range.Value2 = $csv
$excel.DisplayAlerts = $true

# Spaltenüberschriften formatieren
$headerRange = $worksheet.Range("A1:C1")
$headerRange.Font.Bold = $true
$headerRange.Interior.ColorIndex = 15 # Grau

# Filter aktivieren
$worksheet.AutoFilterMode = $true

# Sortierung über Dropdownmenü in Kopfzeilen aktivieren
$headerRange.AutoFilter

# Jede Gruppe farblich markieren
$gruppen = $csv | Select-Object -ExpandProperty Gruppe -Unique
$farben = 10..60 # Farbpalette
$i = 0
foreach ($gruppe in $gruppen) {
    $farbe = $farben[($i++) % $farben.Count]
    $zellen = $worksheet.Cells.Find($gruppe)
    if ($zellen -ne $null) {
        while ($zellen -ne $null) {
            $zellen.EntireRow.Interior.ColorIndex = $farbe
            $zellen = $worksheet.Cells.FindNext($zellen)
        }
    }
}

# Spaltenbreite automatisch anpassen
$worksheet.Columns.AutoFit()

Write-Verbose "Skript abgeschlossen."
