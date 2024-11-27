Überarbeite das folgeden script, optimiere es, nutze wenn möglich vorteile von DFS aus, verbessere die Darstellung in der html


# Pfad (wird vom Benutzer abgefragt)
$pfad = Read-Host -Prompt "Bitte geben Sie den Pfad ein (z.B. \\domain\dfs_namespace)"

# Altersgrenzen in Jahren
$alter30Jahre = 30
$alter10Jahre = 10
$alter6Jahre = 6
$alter5Jahre = 5

# Berechne die Datumsgrenzen
$grenze30Jahre = (Get-Date).AddYears(-$alter30Jahre)
$grenze10Jahre = (Get-Date).AddYears(-$alter10Jahre)
$grenze6Jahre = (Get-Date).AddYears(-$alter6Jahre)
$grenze5Jahre = (Get-Date).AddYears(-$alter5Jahre)

# Definiere die Größenklassen in Bytes
$groesse10MB = 10MB
$groesse100MB = 100MB
$groesse1GB = 1GB

# Ausgabeverzeichnis (Standard: Speicherort des Skripts)
$ausgabeVerzeichnis = "C:\Daten\maybe_unneeded_files"


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

# --- Dateien nach Alter ---

# Erstelle die Baumstrukturen für alte Dateien
ErstelleBaumstruktur -Pfad $pfad -DateiName "AelterAls30Jahre" -Grenzdatum $grenze30Jahre
ErstelleBaumstruktur -Pfad $pfad -DateiName "AelterAls10Jahre" -Grenzdatum $grenze10Jahre
ErstelleBaumstruktur -Pfad $pfad -DateiName "AelterAls6Jahre" -Grenzdatum $grenze6Jahre
ErstelleBaumstruktur -Pfad $pfad -DateiName "AelterAls5Jahre" -Grenzdatum $grenze5Jahre

# Entferne Duplikate aus den Alterslisten, beginnend mit der ältesten Klasse
#Get-Content "$ausgabeVerzeichnis\AelterAls30Jahre.txt" | Select-Object -Unique | Out-File "$ausgabeVerzeichnis\AelterAls30Jahre.txt"
#Get-Content "$ausgabeVerzeichnis\AelterAls10Jahre.txt" | Where-Object {$_ -notin (Get-Content "$ausgabeVerzeichnis\AelterAls30Jahre.txt")} | Select-Object -Unique | Out-File "$ausgabeVerzeichnis\AelterAls10Jahre.txt"
#Get-Content "$ausgabeVerzeichnis\AelterAls6Jahre.txt" | Where-Object {$_ -notin (Get-Content "$ausgabeVerzeichnis\AelterAls30Jahre.txt") -and $_ -notin (Get-Content "$ausgabeVerzeichnis\AelterAls10Jahre.txt")} | Select-Object -Unique | Out-File "$ausgabeVerzeichnis\AelterAls6Jahre.txt"
#Get-Content "$ausgabeVerzeichnis\AelterAls5Jahre.txt" | Where-Object {$_ -notin (Get-Content "$ausgabeVerzeichnis\AelterAls30Jahre.txt") -and $_ -notin (Get-Content "$ausgabeVerzeichnis\AelterAls10Jahre.txt") -and $_ -notin (Get-Content "$ausgabeVerzeichnis\AelterAls6Jahre.txt")} | Select-Object -Unique | Out-File "$ausgabeVerzeichnis\AelterAls5Jahre.txt"

# --- Dateien nach Größe ---

# Erstelle die Baumstrukturen für grosse Dateien
ErstelleBaumstruktur -Pfad $pfad -DateiName "GroesserAls10MB" -Grenzwert $groesse10MB
ErstelleBaumstruktur -Pfad $pfad -DateiName "GroesserAls100MB" -Grenzwert $groesse100MB
ErstelleBaumstruktur -Pfad $pfad -DateiName "GroesserAls1GB" -Grenzwert $groesse1GB

# Entferne Duplikate aus den Grössenlisten, beginnend mit der kleinsten Klasse
#Get-Content "$ausgabeVerzeichnis\GroesserAls10MB.txt" | Select-Object -Unique | Out-File "$ausgabeVerzeichnis\GroesserAls10MB.txt"
#Get-Content "$ausgabeVerzeichnis\GroesserAls100MB.txt" | Where-Object {$_ -notin (Get-Content "$ausgabeVerzeichnis\GroesserAls10MB.txt")} | Select-Object -Unique | Out-File "$ausgabeVerzeichnis\GroesserAls100MB.txt"
#Get-Content "$ausgabeVerzeichnis\GroesserAls1GB.txt" | Where-Object {$_ -notin (Get-Content "$ausgabeVerzeichnis\GroesserAls10MB.txt") -and $_ -notin (Get-Content "$ausgabeVerzeichnis\GroesserAls100MB.txt")} | Select-Object -Unique | Out-File "$ausgabeVerzeichnis\GroesserAls1GB.txt"

# --- Duplikate ---

# Duplikatsuche
Get-ChildItem -Path $pfad -Recurse | Group-Object -Property Name | Where-Object {$_.Count -gt 1} | ForEach-Object {
  $_.Group | Select-Object FullName, Length, LastWriteTime | Out-File -FilePath "$ausgabeVerzeichnis\Duplikate.txt" -Append
}

# --- HTML-Generierung ---

# Funktion zum Erstellen einer HTML-Datei mit visueller Darstellung als Baumstruktur und Kontrollkästchen
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
  ul { list-style-type: none; }
  li { margin-bottom: 5px; }
  .folder { font-weight: bold; cursor: pointer; }
  .file { margin-left: 20px; }
  .checkbox-column { 
    float: left; 
    width: 50px; 
    margin-right: 10px; 
  }
</style>
</head>
<body>
<h1>Dateien - $Titel</h1>
<h2>Pfad: $pfad</h2>
<h2>Datum: $(Get-Date)</h2>
<button onclick="erstelleArchivierungsskript()">Archivierungsskript erstellen</button>
<ul id="tree">
"@

  $dateien = Get-Content -Path "$ausgabeVerzeichnis\$DateiName.txt" | ForEach-Object {$_.Replace("\\", "/")}
  $baumstruktur = ErstelleBaumstrukturAusPfaden($dateien)
  $html += $baumstruktur

  $html += @"
</ul>
<script>
  function erstelleArchivierungsskript() {
    const checkboxes = document.querySelectorAll('input[type="checkbox"]:checked');
    let scriptInhalt = '# Archivierungsskript\n\n';
    checkboxes.forEach(checkbox => {
      const pfad = checkbox.value;
      const zielordner = checkbox.name === 'archiv' ? '---ARCHIV---' : '---ARCHIV-LOCAL---';
      scriptInhalt += `robocopy "${pfad}" "${pfad.substring(0, pfad.lastIndexOf('/'))}/${zielordner}" /MOV /E /LOG:"$ausgabeVerzeichnis\\Archivierungslog.txt"\n`;
    });
    navigator.clipboard.writeText(scriptInhalt)
      .then(() => alert('Das Archivierungsskript wurde in die Zwischenablage kopiert.'))
      .catch(err => console.error('Fehler beim Kopieren in die Zwischenablage: ', err));
  }

  const folders = document.querySelectorAll('.folder');
  folders.forEach(folder => {
    folder.addEventListener('click', () => {
      const files = folder.nextElementSibling;
      if (files.style.display === 'none') {
        files.style.display = 'block';
      } else {
        files.style.display = 'none';
      }
    });
  });
</script>
</body>
</html>
"@

  $html | Out-File -FilePath "$ausgabeVerzeichnis\$DateiName.html"
}

# Funktion zum Erstellen der Baumstruktur aus den Dateipfaden
function ErstelleBaumstrukturAusPfaden {
  param(
    [string[]]$Pfade
  )
  $baum = @{}
  foreach ($pfad in $Pfade) {
    $teile = $pfad.Split("/")
    $aktuellerKnoten = $baum
    for ($i = 0; $i -lt $teile.Length - 1; $i++) {
      $teil = $teile[$i]
      if (-not $aktuellerKnoten.ContainsKey($teil)) {
        $aktuellerKnoten[$teil] = @{}
      }
      $aktuellerKnoten = $aktuellerKnoten[$teil]
    }
    $aktuellerKnoten[$teile[-1]] = $null
  }
  return ErstelleHTMLBaumstruktur($baum)
}

# Funktion zum Erstellen der HTML-Baumstruktur mit Kontrollkästchen
function ErstelleHTMLBaumstruktur {
  param(
    [hashtable]$Baum,
    [int]$Ebene = 0,
    [string]$Pfad = ""
  )
  $html = ""
  foreach ($knoten in $Baum.Keys) {
    $einzug = "  " * $Ebene
    $aktuellerPfad = Join-Path $Pfad $knoten
    if ($Baum[$knoten] -is [hashtable]) {
      $html += "$einzug<li><div class='checkbox-column'><input type='checkbox' name='archiv' value='$aktuellerPfad'> <input type='checkbox' name='archiv-local' value='$aktuellerPfad'></div><div class='folder'>$knoten</div></li>`n"
      $html += "$einzug<ul style='display: none;'>`n"
      $html += ErstelleHTMLBaumstruktur -Baum $Baum[$knoten] -Ebene ($Ebene + 1) -Pfad $aktuellerPfad
      $html += "$einzug</ul>`n"
    } else {
      $html += "$einzug<li><div class='checkbox-column'><input type='checkbox' name='archiv' value='$aktuellerPfad'> <input type='checkbox' name='archiv-local' value='$aktuellerPfad'></div><div class='file'>$knoten</div></li>`n"
    }
  }
  return $html
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

