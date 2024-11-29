# --- Benutzerdefinierte Einstellungen ---
# Altersgrenzen in Jahren
$alter30Jahre = 30
$alter10Jahre = 10
$alter6Jahre = 6
$alter5Jahre = 5

# --- Ende der benutzerdefinierten Einstellungen ---

# Berechne die Datumsgrenzen
$grenze30Jahre = (Get-Date).AddYears(-$alter30Jahre)
$grenze10Jahre = (Get-Date).AddYears(-$alter10Jahre)
$grenze6Jahre = (Get-Date).AddYears(-$alter6Jahre)
$grenze5Jahre = (Get-Date).AddYears(-$alter5Jahre)

# Definiere die Größenklassen in Bytes
$groesse10MB = 10MB
$groesse100MB = 100MB
$groesse1GB = 1GB

# Ausgabeverzeichnis (aktuell genutzter Pfad)
$ausgabeVerzeichnis = d:\daten\maybe_unneded_files

# Funktion zum Erstellen der Baumstruktur und Speichern in einer Textdatei
function ErstelleBaumstruktur {
  param(
    [string]$Pfad,
    [string]$DateiName,
    [datetime]$Grenzdatum,
    [long]$Grenzwert
  )
  if ($Grenzdatum) {
    Get-ChildItem -Path $Pfad -Recurse | Where-Object { $_.LastWriteTime -lt $Grenzdatum } | ForEach-Object {
      $_.FullName | Out-File -FilePath "$ausgabeVerzeichnis\$DateiName.txt" -Append
    }
  }
  elseif ($Grenzwert) {
    Get-ChildItem -Path $Pfad -Recurse | Where-Object { $_.Length -gt $Grenzwert } | ForEach-Object {
      $_.FullName | Out-File -FilePath "$ausgabeVerzeichnis\$DateiName.txt" -Append
    }
  }
}

# Erstelle die Baumstrukturen für alte Dateien
ErstelleBaumstruktur -Pfad "." -DateiName "AelterAls30Jahre" -Grenzdatum $grenze30Jahre
ErstelleBaumstruktur -Pfad "." -DateiName "AelterAls10Jahre" -Grenzdatum $grenze10Jahre
ErstelleBaumstruktur -Pfad "." -DateiName "AelterAls6Jahre" -Grenzdatum $grenze6Jahre
ErstelleBaumstruktur -Pfad "." -DateiName "AelterAls5Jahre" -Grenzdatum $grenze5Jahre

# Erstelle die Baumstrukturen für grosse Dateien
ErstelleBaumstruktur -Pfad "." -DateiName "GroesserAls10MB" -Grenzwert $groesse10MB
ErstelleBaumstruktur -Pfad "." -DateiName "GroesserAls100MB" -Grenzwert $groesse100MB
ErstelleBaumstruktur -Pfad "." -DateiName "GroesserAls1GB" -Grenzwert $groesse1GB

# Entferne Duplikate aus den Grössenlisten
Get-Content "$ausgabeVerzeichnis\GroesserAls10MB.txt" | Select-Object -Unique | Out-File "$ausgabeVerzeichnis\GroesserAls10MB.txt"
Get-Content "$ausgabeVerzeichnis\GroesserAls100MB.txt" | Where-Object {$_ -notin (Get-Content "$ausgabeVerzeichnis\GroesserAls10MB.txt")} | Select-Object -Unique | Out-File "$ausgabeVerzeichnis\GroesserAls100MB.txt"
Get-Content "$ausgabeVerzeichnis\GroesserAls1GB.txt" | Where-Object {$_ -notin (Get-Content "$ausgabeVerzeichnis\GroesserAls10MB.txt") -and $_ -notin (Get-Content "$ausgabeVerzeichnis\GroesserAls100MB.txt")} | Select-Object -Unique | Out-File "$ausgabeVerzeichnis\GroesserAls1GB.txt"

# Duplikatsuche
Get-ChildItem -Path "." -Recurse | Group-Object -Property Name | Where-Object {$_.Count -gt 1} | ForEach-Object {
  $_.Group | Select-Object FullName, Length, LastWriteTime | Out-File -FilePath "$ausgabeVerzeichnis\Duplikate.txt" -Append
}

# Funktion zum Erstellen einer HTML-Datei mit Dark Mode UI
function ErstelleHTML {
  param(
    [string]$DateiName,
    [string]$Titel
  )
  $html = @"
<!DOCTYPE html>
<html>
<head>
<title>Dateien - $Titel</title>
<style>
  body {
    background-color: #121212;
    color: #FFFFFF;
    font-family: Arial, sans-serif;
  }
  h1 {
    color: #BB86FC;
  }
  ul {
    list-style-type: none;
    padding: 0;
  }
  li {
    margin: 5px 0;
  }
  li::before {
    content: '📁 ';
  }
</style>
</head>
<body>
<h1>Dateien - $Titel</h1>
<ul>
"@

  Get-Content -Path ".\$DateiName.txt" | ForEach-Object {
    $pfad = $_ -replace "\\", "/"
    $html += "<li>$pfad</li>"
  }

  $html += @"
</ul>
</body>
</html>
"@

  $html | Out-File -FilePath "$ausgabeVerzeichnis\$DateiName.html"
}

# Erstelle die HTML-Dateien
ErstelleHTML -DateiName "AelterAls30Jahre" -Titel "Älter als 30 Jahre"
ErstelleHTML -DateiName "AelterAls10Jahre" -Titel "Älter als 10 Jahre"
ErstelleHTML -DateiName "AelterAls6Jahre" -Titel "Älter als 6 Jahre"
ErstelleHTML -DateiName "AelterAls5Jahre" -Titel "Älter als 5 Jahre"
ErstelleHTML -DateiName "GroesserAls10MB" -Titel "Größer als 10 MB"
ErstelleHTML -DateiName "GroesserAls100MB" -Titel "Größer als 100 MB"
ErstelleHTML -DateiName "GroesserAls1GB" -Titel "Größer als 1 GB"
ErstelleHTML -DateiName "Duplikate" -Titel "Duplikate"

Write-Host "Skript abgeschlossen. Die Dateien befinden sich in: $ausgabeVerzeichnis"
