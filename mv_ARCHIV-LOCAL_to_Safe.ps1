# Standardpfade
$defaultSmbPfad = "\\server\freigabe"
$defaultZielPfad = "e:\Daten\Archiv"

# Abfrage des SMB Serverpfads mit Standardpfad
$smbPfad = Read-Host -Prompt "Geben Sie den SMB Serverpfad ein (Standard: $defaultSmbPfad)"
if ([string]::IsNullOrEmpty($smbPfad)) {
  $smbPfad = $defaultSmbPfad
}

# Abfrage des Zielpfads mit Standardpfad
$zielPfad = Read-Host -Prompt "Geben Sie den Zielpfad ein (Standard: $defaultZielPfad)"
if ([string]::IsNullOrEmpty($zielPfad)) {
  $zielPfad = $defaultZielPfad
}

# Überprüfung, ob die Pfade existieren
if (!(Test-Path $smbPfad)) {
  Write-Error "Der angegebene SMB Pfad existiert nicht: $smbPfad"
  exit 1
}
if (!(Test-Path $zielPfad)) {
  Write-Error "Der angegebene Zielpfad existiert nicht: $zielPfad"
  exit 1
}

# Protokolldatei erstellen
$protokollDatei = Join-Path $zielPfad "archiv_protokoll.txt"
"Protokoll gestartet am $(Get-Date)" | Out-File -FilePath $protokollDatei -Encoding UTF8

# Schleife durch alle Unterordner im SMB Pfad
Get-ChildItem -Path $smbPfad -Directory | ForEach-Object {
  $ordnerName = $_.Name
  $archivOrdner = Join-Path $_.FullName "---archive-Local---"

  # Überprüfung, ob der Archivordner existiert
  if (Test-Path $archivOrdner) {
    $zielOrdner = Join-Path $zielPfad $ordnerName

    # Erstelle den Zielordner, falls er nicht existiert
    if (!(Test-Path $zielOrdner)) {
      New-Item -Path $zielOrdner -ItemType Directory | Out-Null
    }

    # Verschiebe den Archivordner mit Robocopy
    robocopy $archivOrdner $zielOrdner /MOVE /E /R:3 /W:5 /LOG+:$protokollDatei
    "Archivordner '$archivOrdner' wurde nach '$zielOrdner' verschoben." | Out-File -FilePath $protokollDatei -Encoding UTF8 -Append
  }
}

"Skript abgeschlossen am $(Get-Date)" | Out-File -FilePath $protokollDatei -Encoding UTF8 -Append
Write-Host "Skript abgeschlossen. Protokolldatei: $protokollDatei"
