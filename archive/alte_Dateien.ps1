# Definiere den DFS-Pfad und die Altersgrenzen
$dfsPfad = "\\domain\dfs_namespace"  # Ersetze dies mit dem tatsächlichen DFS-Pfad
$alter30Jahre = 30
$alter10Jahre = 10
$alter6Jahre = 6
$alter5Jahre = 5

# Berechne die Datumsgrenzen
$grenze30Jahre = (Get-Date).AddYears(-$alter30Jahre)
$grenze10Jahre = (Get-Date).AddYears(-$alter10Jahre)
$grenze6Jahre = (Get-Date).AddYears(-$alter6Jahre)
$grenze5Jahre = (Get-Date).AddYears(-$alter5Jahre)

# Erstelle die Ausgabeverzeichnisse, falls sie nicht existieren
$ausgabeVerzeichnis = "C:\DFS_Dateien"  # Passe den Pfad an, falls erforderlich
New-Item -ItemType Directory -Path $ausgabeVerzeichnis -Force -ErrorAction SilentlyContinue

# Funktion zum Erstellen der Baumstruktur und Speichern in einer Textdatei
function ErstelleBaumstruktur {
    param(
        [string]$Pfad,
        [string]$DateiName,
        [datetime]$Grenzdatum
    )
    Get-ChildItem -Path $Pfad -Recurse | Where-Object { $_.LastWriteTime -lt $Grenzdatum } | ForEach-Object {
        $_.FullName | Out-File -FilePath "$ausgabeVerzeichnis\$DateiName.txt" -Append
    }
}

# Erstelle die Baumstrukturen und speichere sie in Textdateien
ErstelleBaumstruktur -Pfad $dfsPfad -DateiName "AelterAls30Jahre" -Grenzdatum $grenze30Jahre
ErstelleBaumstruktur -Pfad $dfsPfad -DateiName "AelterAls10Jahre" -Grenzdatum $grenze10Jahre
ErstelleBaumstruktur -Pfad $dfsPfad -DateiName "AelterAls6Jahre" -Grenzdatum $grenze6Jahre
ErstelleBaumstruktur -Pfad $dfsPfad -DateiName "AelterAls5Jahre" -Grenzdatum $grenze5Jahre

# Funktion zum Erstellen einer HTML-Datei mit visueller Darstellung
function ErstelleHTML {
    param(
        [string]$DateiName
    )
    $html = @"
<!DOCTYPE html>
<html>
<head>
<title>DFS Dateien - $DateiName</title>
</head>
<body>
<h1>DFS Dateien - $DateiName</h1>
<ul>
"@

    Get-Content -Path "$ausgabeVerzeichnis\$DateiName.txt" | ForEach-Object {
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
ErstelleHTML -DateiName "AelterAls30Jahre"
ErstelleHTML -DateiName "AelterAls10Jahre"
ErstelleHTML -DateiName "AelterAls6Jahre"
ErstelleHTML -DateiName "AelterAls5Jahre"

Write-Host "Skript abgeschlossen. Die Dateien befinden sich in: $ausgabeVerzeichnis"
