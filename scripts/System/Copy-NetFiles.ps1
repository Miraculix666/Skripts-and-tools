# Konfiguration
$remoteComputer = ""  # Wird im Skript abgefragt
$usbDrive = "E:"
$localFolder = "C:\Daten\move"
$networkDrive = "Z:\Zielordner"
$folderToCopy = "move"
$logFileName = "kopierlog.txt"

# Funktion zum Anzeigen des Menüs
function Show-Menu {
    Clear-Host
    if ($remoteComputer -ne "") {
        Write-Host "Verbunden mit: $remoteComputer" -ForegroundColor Green
    }
    Write-Host "========================================"
    Write-Host "             Kopiermenü"
    Write-Host "========================================"
    Write-Host "1: Vom USB-Stick auf den lokalen Rechner"
    Write-Host "2: Vom lokalen Rechner auf den USB-Stick"
    Write-Host "3: Vom USB-Stick auf das Netzlaufwerk"
    Write-Host "4: Vom Netzlaufwerk auf den USB-Stick"
    Write-Host "5: Vom lokalen Rechner auf das Netzlaufwerk"
    Write-Host "6: Vom Netzlaufwerk auf den lokalen Rechner"
    Write-Host "7: Anderen Remote-Computer auswählen"
    Write-Host "0: Beenden"
    Write-Host ""
}

# Funktion zum Schreiben von Log-Einträgen
function Write-Log {
    param(
        [string]$message,
        [string]$logFilePath,
        [ConsoleColor]$color = "White"  # Standardfarbe ist Weiß
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] $message"

    # In die Konsole schreiben mit der angegebenen Farbe
    Write-Host $logEntry -ForegroundColor $color

    # In die Logdatei schreiben
    try {
        Out-File -FilePath $logFilePath -InputObject $logEntry -Append -ErrorAction Stop
    } catch {
        Write-Host "Fehler beim Schreiben in die Logdatei: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Funktion zum Kopieren von Dateien mit ausführlicher Ausgabe und Logging
function Copy-Files {
    param(
        [string]$source,
        [string]$destination,
        [string]$sourceName,
        [string]$destinationName
    )

    $logFilePath = Join-Path -Path $destination -ChildPath $logFileName
    Write-Log -message "Starte Kopiervorgang von '$sourceName' nach '$destinationName'..." -logFilePath $logFilePath -color Yellow

    # Prüfe ob der Quellpfad existiert
    if (!(Test-Path $source)) {
        Write-Log -message "Quelle '$source' existiert nicht." -logFilePath $logFilePath -color Red
        return
    }

    # Prüfe ob das Ziel ein Verzeichnis ist, falls nicht, versuche es zu erstellen
    if (!(Test-Path $destination)) {
        Write-Log -message "Zielverzeichnis '$destination' existiert nicht, versuche es zu erstellen..." -logFilePath $logFilePath
        try {
            New-Item -ItemType Directory -Path $destination -ErrorAction Stop | Out-Null
            Write-Log -message "Zielverzeichnis '$destination' erfolgreich erstellt." -logFilePath $logFilePath -color Green
        } catch {
            Write-Log -message "Fehler beim Erstellen des Zielverzeichnisses '$destination': $($_.Exception.Message)" -logFilePath $logFilePath -color Red
            return
        }
    }

    try {
        Copy-Item -Path $source -Destination $destination -Recurse -Force -Verbose 4>&1 | ForEach-Object {
            # Verarbeite die Verbose-Ausgabe und logge den Fortschritt
            if ($_.CategoryInfo.Category -eq "CopyFile") {
                Write-Log -message $_.Target -logFilePath $logFilePath -color Green
            } else {
                Write-Log -message $_ -logFilePath $logFilePath
            }
        }
        Write-Log -message "Kopiervorgang erfolgreich abgeschlossen." -logFilePath $logFilePath -color Green
    } catch {
        Write-Log -message "Fehler beim Kopieren: $($_.Exception.Message)" -logFilePath $logFilePath -color Red
    }
}

# Hauptteil des Skripts
# Verbindungsinformationen abfragen, falls noch nicht gesetzt
if ($remoteComputer -eq "") {
    $remoteComputer = Read-Host "Gib den Namen des Remote-Computers ein"
}

# Menüschleife
while ($true) {
    Show-Menu
    $choice = Read-Host "Wähle eine Option"

    switch ($choice) {
        "1" {
            # Vom USB-Stick auf den lokalen Rechner
            $source = Join-Path -Path $usbDrive -ChildPath $folderToCopy
            $destination = $localFolder
            Copy-Files -source $source -destination $destination -sourceName "USB-Stick ($source)" -destinationName "Lokaler Rechner ($destination)"
        }
        "2" {
            # Vom lokalen Rechner auf den USB-Stick
            $source = $localFolder
            $destination = Join-Path -Path $usbDrive -ChildPath $folderToCopy
            Copy-Files -source $source -destination $destination -sourceName "Lokaler Rechner ($source)" -destinationName "USB-Stick ($destination)"
        }
        "3" {
            # Vom USB-Stick auf das Netzlaufwerk
            $source = Join-Path -Path $usbDrive -ChildPath $folderToCopy
            $destination = $networkDrive
            Copy-Files -source $source -destination $destination -sourceName "USB-Stick ($source)" -destinationName "Netzlaufwerk ($destination)"
        }
        "4" {
            # Vom Netzlaufwerk auf den USB-Stick
            $source = $networkDrive
            $destination = Join-Path -Path $usbDrive -ChildPath $folderToCopy
            Copy-Files -source $source -destination $destination -sourceName "Netzlaufwerk ($source)" -destinationName "USB-Stick ($destination)"
        }
        "5" {
            # Vom lokalen Rechner auf das Netzlaufwerk
            $source = $localFolder
            $destination = $networkDrive
            Copy-Files -source $source -destination $destination -sourceName "Lokaler Rechner ($source)" -destinationName "Netzlaufwerk ($destination)"
        }
        "6" {
            # Vom Netzlaufwerk auf den lokalen Rechner
            $source = $networkDrive
            $destination = $localFolder
            Copy-Files -source $source -destination $destination -sourceName "Netzlaufwerk ($source)" -destinationName "Lokaler Rechner ($destination)"
        }
        "7" {
            # Anderen Remote-Computer auswählen
            $remoteComputer = Read-Host "Gib den Namen des neuen Remote-Computers ein"
        }
        "0" {
            # Beenden
            exit
        }
        default {
            Write-Host "Ungültige Auswahl. Bitte wähle eine Zahl von 0 bis 7." -ForegroundColor Red
        }
    }

    Read-Host "Drücke Enter, um fortzufahren..."
}
