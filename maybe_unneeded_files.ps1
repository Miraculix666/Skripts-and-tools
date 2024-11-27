# Pfad (wird vom Benutzer abgefragt)
$pfad = Read-Host -Prompt "Bitte geben Sie den Pfad ein (z.B. \\domain\dfs_namespace)"

# Altersgrenzen und Größenklassen
$alterGrenzen = @{30 = (Get-Date).AddYears(-30); 10 = (Get-Date).AddYears(-10); 6 = (Get-Date).AddYears(-6); 5 = (Get-Date).AddYears(-5)}
$groessenKlassen = @{10MB = 10MB; 100MB = 100MB; 1GB = 1GB}

# Ausgabeverzeichnis
$ausgabeVerzeichnis = "C:\Daten\maybe_unneeded_files"
New-Item -ItemType Directory -Force -Path $ausgabeVerzeichnis | Out-Null

# Funktion zum Erstellen der Baumstruktur
function Get-Baumstruktur {
    param(
        [string]$Pfad,
        [string]$DateiName,
        [datetime]$Grenzdatum,
        [long]$Grenzwert
    )

    $filter = if ($Grenzdatum) { {$_.LastWriteTime -lt $Grenzdatum} } else { {$_.Length -gt $Grenzwert} }
    
    Get-ChildItem -Path $Pfad -Recurse -File | Where-Object $filter | Select-Object -ExpandProperty FullName |
    Out-File -FilePath "$ausgabeVerzeichnis\$DateiName.txt"
}

# Dateien nach Alter und Größe erfassen
foreach ($alter in $alterGrenzen.Keys) {
    Get-Baumstruktur -Pfad $pfad -DateiName "AelterAls${alter}Jahre" -Grenzdatum $alterGrenzen[$alter]
}

foreach ($groesse in $groessenKlassen.Keys) {
    Get-Baumstruktur -Pfad $pfad -DateiName "GroesserAls$groesse" -Grenzwert $groessenKlassen[$groesse]
}

# Duplikate finden
Get-ChildItem -Path $pfad -Recurse -File | Group-Object -Property Name | Where-Object {$_.Count -gt 1} | 
ForEach-Object { $_.Group | Select-Object FullName, Length, LastWriteTime } | 
Export-Csv -Path "$ausgabeVerzeichnis\Duplikate.csv" -NoTypeInformation

# HTML-Generierung
function New-HTMLDarstellung {
    param(
        [string]$DateiName,
        [string]$Titel
    )

    $dateien = Get-Content -Path "$ausgabeVerzeichnis\$DateiName.txt" | ForEach-Object {$_.Replace("\", "/")}
    $baumstruktur = New-BaumstrukturAusPfaden $dateien

    $html = @"
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dateien - $Titel</title>
    <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
        ul { list-style-type: none; padding-left: 20px; }
        li { margin-bottom: 5px; }
        .folder { font-weight: bold; cursor: pointer; }
        .file { margin-left: 20px; }
        .checkbox-column { display: inline-block; width: 50px; }
        .toggle-all { margin-bottom: 10px; }
    </style>
</head>
<body>
    <h1>Dateien - $Titel</h1>
    <h2>Pfad: $pfad</h2>
    <h2>Datum: $(Get-Date -Format "dd.MM.yyyy HH:mm")</h2>
    <button onclick="erstelleArchivierungsskript()">Archivierungsskript erstellen</button>
    <div class="toggle-all">
        <button onclick="toggleAll(true)">Alle ausklappen</button>
        <button onclick="toggleAll(false)">Alle einklappen</button>
    </div>
    <ul id="tree">
    $baumstruktur
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

    function toggleFolder(element) {
        const files = element.nextElementSibling;
        files.style.display = files.style.display === 'none' ? 'block' : 'none';
    }

    function toggleAll(expand) {
        document.querySelectorAll('ul').forEach(ul => {
            ul.style.display = expand ? 'block' : 'none';
        });
    }

    document.querySelectorAll('.folder').forEach(folder => {
        folder.addEventListener('click', () => toggleFolder(folder));
    });
    </script>
</body>
</html>
"@

    $html | Out-File -FilePath "$ausgabeVerzeichnis\$DateiName.html"
}

function New-BaumstrukturAusPfaden {
    param([string[]]$Pfade)

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
    return (New-HTMLBaumstruktur $baum)
}

function New-HTMLBaumstruktur {
    param(
        [hashtable]$Baum,
        [int]$Ebene = 0,
        [string]$Pfad = ""
    )

    $html = ""
    foreach ($knoten in $Baum.Keys) {
        $einzug = "    " * $Ebene
        $aktuellerPfad = Join-Path $Pfad $knoten -Replace "\\", "/"
        if ($Baum[$knoten] -is [hashtable]) {
            $html += "$einzug<li><div class='checkbox-column'><input type='checkbox' name='archiv' value='$aktuellerPfad'> <input type='checkbox' name='archiv-local' value='$aktuellerPfad'></div><span class='folder'>$knoten</span></li>`n"
            $html += "$einzug<ul style='display: none;'>`n"
            $html += New-HTMLBaumstruktur -Baum $Baum[$knoten] -Ebene ($Ebene + 1) -Pfad $aktuellerPfad
            $html += "$einzug</ul>`n"
        } else {
            $html += "$einzug<li><div class='checkbox-column'><input type='checkbox' name='archiv' value='$aktuellerPfad'> <input type='checkbox' name='archiv-local' value='$aktuellerPfad'></div><span class='file'>$knoten</span></li>`n"
        }
    }
    return $html
}

# HTML-Dateien erstellen
@("AelterAls30Jahre", "AelterAls10Jahre", "AelterAls6Jahre", "AelterAls5Jahre", "GroesserAls10MB", "GroesserAls100MB", "GroesserAls1GB") | ForEach-Object {
    New-HTMLDarstellung -DateiName $_ -Titel $_.Replace("AelterAls", "Älter als ").Replace("GroesserAls", "Größer als ")
}

New-HTMLDarstellung -DateiName "Duplikate" -Titel "Duplikate"

Write-Host "Skript abgeschlossen. Die Dateien befinden sich in: $ausgabeVerzeichnis"
