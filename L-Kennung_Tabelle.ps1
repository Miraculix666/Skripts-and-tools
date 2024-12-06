# Benutzer mit Anmeldenamen beginnend mit "L" abrufen
Write-Verbose "Suche nach AD-Benutzern mit Anmeldenamen beginnend mit 'L'..."
$benutzer = Get-ADUser -Filter "SamAccountName -like 'L*'" -Properties SamAccountName, Name, MemberOf, DistinguishedName

# Ausgabe der gefundenen Benutzer in der Shell
Write-Host "Gefundene Benutzer:"
$benutzer | Format-Table SamAccountName, Name

# Benutzerinformationen und Gruppenzugehörigkeiten in einer Tabelle zusammenfassen
Write-Verbose "Erstelle Tabelle mit Benutzerinformationen und Gruppenzugehörigkeiten..."
$tabelle = @()
$benutzer | ForEach-Object {
    Write-Verbose "Verarbeite Benutzer '$($_.SamAccountName)'..."
    $ou = $_.DistinguishedName -replace '^CN=[^,]+,OU=([^,]+),.*$', '$1'
    # OU "Benutzer" ignorieren
    if ($ou -ne "Benutzer") {
        $gruppen = $_.MemberOf | Get-ADGroup | Select-Object -ExpandProperty Name
        $gruppenHash = @{}
        foreach ($gruppe in $gruppen) {
            $gruppenHash[$gruppe] = $gruppe
        }
        $tabelle += [PSCustomObject]@{
            OU = $ou
            Benutzer = $_.Name
            Gruppen = $gruppenHash
        }
        # Ausgabe der aktuellen Benutzer-Gruppen-Zuordnung in der Shell
        Write-Host "Benutzer: $($_.Name), OU: $ou, Gruppen: $($gruppen -join ', ')"
    }
}

# Alle Gruppennamen erfassen und als Spalten in der Tabelle hinzufügen
$alleGruppen = $tabelle | ForEach-Object { $_.Gruppen.Keys } | Select-Object -Unique
$tabelle | ForEach-Object {
    foreach ($gruppe in $alleGruppen) {
        if (-not $_.PSObject.Properties.Match($gruppe)) {
            Add-Member -InputObject $_ -MemberType NoteProperty -Name $gruppe -Value ($_.Gruppen[$gruppe] -ne $null)
        }
    }
}

# Tabelle nach OU, Benutzer und dann nach Gruppen sortieren
Write-Verbose "Sortiere Tabelle nach OU, Benutzer und Gruppen..."
$tabelle = $tabelle | Sort-Object -Property OU, Benutzer

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
$headerRange = $worksheet.Range("A1:Z1") # Passen Sie den Bereich entsprechend der Anzahl der Spalten an
$headerRange.Font.Bold = $true
$headerRange.Interior.ColorIndex = 15 # Grau

# Filter aktivieren
$worksheet.AutoFilterMode = $false # Deaktiviert den AutoFilter-Modus vor dem Aktivieren
$headerRange
