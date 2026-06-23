<#
.SYNOPSIS
Ueberwacht einen Ordner auf neue Dateien und versendet diese automatisch als E-Mail-Anhang.

.DESCRIPTION
Dieses PowerShell-Skript ueberwacht ein konfiguriertes Zielverzeichnis. Neue Dateien, die noch nicht per E-Mail versendet wurden,
werden automatisch als Anhang ueber einen SMTP-Server gesendet. Der Versand wird in einer Protokolldatei erfasst, um doppelte
E-Mails zu verhindern. Alternativ koennen versendete Dateien in einen Archivordner verschoben werden.

.PARAMETER FolderPath
Der Pfad des Ordners, der auf neue Dateien ueberwacht werden soll.

.PARAMETER ToEmail
Die E-Mail-Adresse des Empfaengers.

.PARAMETER FromEmail
Die E-Mail-Adresse des Absenders.

.PARAMETER SmtpServer
Der SMTP-Server, der zum Versenden der E-Mails verwendet werden soll (z.B. "smtp.example.com").

.PARAMETER SmtpPort
Der Port des SMTP-Servers. Standardwert ist 25 (oder 587 fuer TLS/SSL).

.PARAMETER UseSSL
Aktiviert SSL/TLS fuer die SMTP-Verbindung.

.PARAMETER Credential
Optionale Anmeldedaten fuer den SMTP-Server.

.PARAMETER LogFilePath
Pfad zur Protokolldatei, in der bereits versendete Dateinamen aufgezeichnet werden.

.PARAMETER ArchiveFolder
Optionale Angabe eines Archivordners. Wenn angegeben, werden versendete Dateien dorthin verschoben.

.PARAMETER Daemon
Führt das Skript als dauerhaften Dienst (Endlosschleife) aus. Standardmaessig prueft das Skript den Ordner nur einmal und beendet sich dann (perfekt fuer geplante Aufgaben / Task Scheduler).

.PARAMETER IntervalSeconds
Das Intervall in Sekunden, in dem das Verzeichnis im Daemon-Modus ueberprueft wird. Standardwert ist 3600 Sekunden (1 Stunde).

.EXAMPLE
.\Send-NewFilesMail.ps1 -FolderPath "C:\Bilder" -ToEmail "empfaenger@test.de" -FromEmail "sender@test.de" -SmtpServer "mail.test.de"
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$FolderPath,

    [Parameter(Mandatory = $true)]
    [string]$ToEmail,

    [Parameter(Mandatory = $true)]
    [string]$FromEmail,

    [Parameter(Mandatory = $false)]
    [string]$SmtpServer = "localhost",

    [Parameter(Mandatory = $false)]
    [int]$SmtpPort = 25,

    [Parameter(Mandatory = $false)]
    [switch]$UseSSL,

    [Parameter(Mandatory = $false)]
    [pscredential]$Credential,

    [Parameter(Mandatory = $false)]
    [string]$LogFilePath,

    [Parameter(Mandatory = $false)]
    [string]$ArchiveFolder,

    [Parameter(Mandatory = $false)]
    [switch]$Daemon,

    [Parameter(Mandatory = $false)]
    [int]$IntervalSeconds = 3600
)

# Standardpfad fuer Logdatei setzen, falls nicht angegeben
if (-not $LogFilePath) {
    $LogFilePath = Join-Path $PSScriptRoot "SendMail_Log.txt"
}

# Richte UTF-8 Output-Codierung ein
$OutputEncoding = [System.Text.Encoding]::UTF8

function Send-Email {
    param (
        [string]$From,
        [string]$To,
        [string]$Subject,
        [string]$Body,
        [string]$AttachmentPath
    )

    try {
        # Instanziere .NET MailMessage
        $mail = New-Object System.Net.Mail.MailMessage
        $mail.From = New-Object System.Net.Mail.MailAddress($From)
        $mail.To.Add($To)
        $mail.Subject = $Subject
        $mail.Body = $Body

        # Anhang hinzufuegen, falls vorhanden
        if ($AttachmentPath -and (Test-Path $AttachmentPath)) {
            $attachment = New-Object System.Net.Mail.Attachment($AttachmentPath)
            $mail.Attachments.Add($attachment)
        }

        # Instanziere .NET SmtpClient
        $smtp = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
        
        if ($UseSSL) {
            $smtp.EnableSsl = $true
        }

        if ($Credential) {
            $smtp.Credentials = $Credential.GetNetworkCredential()
        }

        # E-Mail senden
        $smtp.Send($mail)
        
        # Ressourcen freigeben
        if ($attachment) { $attachment.Dispose() }
        $mail.Dispose()
        $smtp.Dispose()

        Write-Host "     [OK] E-Mail erfolgreich gesendet mit Anhang: $(Split-Path $AttachmentPath -Leaf)" -ForegroundColor Green
        return $true
    } catch {
        Write-Host "     [ERR] Fehler beim Versenden der E-Mail: $_" -ForegroundColor Red
        return $false
    }
}

function Check-And-Send {
    if (-not (Test-Path $FolderPath)) {
        Write-Warning "Der angegebene Ordner existiert nicht: $FolderPath"
        return
    }

    # Lese den Inhalt der Logdatei ein
    $sentFiles = @()
    if (Test-Path $LogFilePath) {
        $sentFiles = Get-Content -Path $LogFilePath -ErrorAction SilentlyContinue
    } else {
        # Erstelle leere Logdatei
        New-Item -Path $LogFilePath -ItemType File -Force | Out-Null
    }

    # Suche alle Dateien im Zielordner
    $files = Get-ChildItem -Path $FolderPath -File

    if ($files.Count -eq 0) {
        Write-Host "  - Keine Dateien im Verzeichnis gefunden." -ForegroundColor DarkGray
        return
    }

    Write-Host "[+] Analysiere $($files.Count) Datei(en) in $FolderPath..." -ForegroundColor Cyan

    foreach ($file in $files) {
        # Pruefe, ob die Datei bereits gesendet wurde
        if ($sentFiles -contains $file.Name) {
            Write-Host "  - Datei wurde bereits gesendet: $($file.Name)" -ForegroundColor DarkGray
            continue
        }

        Write-Host "  - Neue Datei gefunden: $($file.Name). Starte Versand..." -ForegroundColor Yellow

        $subject = "Automatische Benachrichtigung: Neue Datei $($file.Name)"
        $body = "Hallo,`n`nes wurde eine neue Datei im ueberwachten Verzeichnis erkannt.`n`nDatei: $($file.Name)`nErstellungszeit: $($file.CreationTime)`nGroesse: $([math]::Round($file.Length / 1KB, 2)) KB`n`nBeste Gruesse`nFile Auto Watcher"

        $success = Send-Email -From $FromEmail -To $ToEmail -Subject $subject -Body $body -AttachmentPath $file.FullName

        if ($success) {
            # Schreibe in Logdatei
            Add-Content -Path $LogFilePath -Value $file.Name
            
            # Verschiebe in Archivordner falls konfiguriert
            if ($ArchiveFolder) {
                if (-not (Test-Path $ArchiveFolder)) {
                    New-Item -Path $ArchiveFolder -ItemType Directory -Force | Out-Null
                }
                try {
                    $targetPath = Join-Path $ArchiveFolder $file.Name
                    Move-Item -Path $file.FullName -Destination $targetPath -Force -ErrorAction Stop
                    Write-Host "     -> Datei ins Archiv verschoben: $targetPath" -ForegroundColor DarkGray
                } catch {
                    Write-Warning "Fehler beim Verschieben ins Archiv: $_"
                }
            }
        }
    }
}

function Main {
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host "          AUTO FILE WATCHER & MAIL SENDER         " -ForegroundColor Cyan
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host "Ueberwachungsordner: $FolderPath" -ForegroundColor Yellow
    Write-Host "Empfaenger:          $ToEmail" -ForegroundColor Yellow
    Write-Host "Startzeit:           $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n" -ForegroundColor Gray

    if ($Daemon) {
        Write-Host "[+] Starte Skript im Daemon-Modus (Endlosschleife)..." -ForegroundColor Cyan
        Write-Host "    Pruefintervall: $IntervalSeconds Sekunden. Druecke Strg+C zum Beenden.`n" -ForegroundColor DarkGray
        while ($true) {
            Check-And-Send
            Start-Sleep -Seconds $IntervalSeconds
        }
    } else {
        Write-Host "[+] Führe einmalige Ueberpruefung durch..." -ForegroundColor Cyan
        Check-And-Send
        Write-Host "`n[OK] Ueberpruefung abgeschlossen." -ForegroundColor Green
    }
    Write-Host "==================================================" -ForegroundColor Cyan
}

# Starte das Hauptprogramm
Main
