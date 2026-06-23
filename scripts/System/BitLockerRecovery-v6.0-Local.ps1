# Start-IntelligentBitLockerRecovery-v6.2-Local-FINAL.ps1 - Finale, stabile Version für die lokale Datenrettung mit Auto-Setup.
#Requires -RunAsAdministrator

<#
.SYNOPSIS
  Eine finale, hochrobuste Lösung zur Rettung von Daten aus beschädigten BitLocker-Laufwerken in eine VHDX-Datei auf einem lokalen Laufwerk (z.B. eine externe USB-Festplatte).
.DESCRIPTION
  Diese Version ist das Endprodukt unserer Analyse und umgeht alle Netzwerk- und Berechtigungsprobleme, indem sie ausschließlich lokal arbeitet.
  Sie prüft und installiert bei Bedarf automatisch die für die Ausführung notwendigen Windows-Module (Hyper-V und BitLocker-Tools).
.NOTES
  Autor: PS-Coding (AI-Assisted)
  Version: 6.2 (Local Recovery - FINAL)
  Änderungen:
  - v6.2: Die Funktion zur automatischen Überprüfung und Installation der Voraussetzungen (Hyper-V und BitLocker-Module) wurde wieder hinzugefügt.
  - v6.0: Komplette Entfernung aller Netzwerklogik. Das Skript ist nun ausschließlich für die Rettung auf ein lokales Ziel (z.B. externe USB-Festplatte) ausgelegt.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter()]
    [string]$SourceDrive,
    [Parameter()]
    [string]$RecoveryKey,
    [Parameter()]
    [string]$DestinationFolderPath,
    [Parameter()]
    [string]$VhdxFileName,
    [Parameter()]
    [switch]$Force
)

begin {
    Set-Culture de-DE | Out-Null
    [System.Threading.Thread]::CurrentThread.CurrentUICulture = 'de-DE'
    $VerbosePreference = 'Continue'

    # --- START: Funktion zur Überprüfung und Installation von Abhängigkeiten ---
    function Ensure-Prerequisites {
        Write-Verbose "Prüfe und installiere erforderliche Komponenten..."

        # 1. Hyper-V Modul (für VHD-Befehle)
        if (-not (Get-Module -ListAvailable -Name Hyper-V)) {
            Write-Warning "PowerShell-Modul 'Hyper-V' wurde nicht gefunden. Versuche automatische Installation..."
            if ($PSCmdlet.ShouldProcess("Windows-Feature 'Microsoft-Hyper-V-Management-PowerShell'", "Installieren")) {
                try {
                    Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V-Management-PowerShell -NoRestart -ErrorAction Stop
                    Write-Host "Hyper-V Management Tools erfolgreich installiert." -ForegroundColor Green
                } catch {
                    throw "Fehler bei der Installation der Hyper-V Tools. Bitte manuell installieren und das Skript neu starten."
                }
            } else {
                throw "Die Installation des Hyper-V Moduls wurde abgelehnt. Das Skript kann nicht fortgesetzt werden."
            }
        } else {
            Write-Verbose "Hyper-V Modul ist bereits vorhanden."
        }

        # 2. BitLocker-Tools (für repair-bde)
        if (-not (Get-Command -Name "repair-bde" -ErrorAction SilentlyContinue)) {
            Write-Warning "Befehl 'repair-bde' nicht gefunden. Versuche Installation der BitLocker RSAT-Tools..."
             if ($PSCmdlet.ShouldProcess("Windows-Capability 'Rsat.BitLocker.Recovery.Tools'", "Installieren")) {
                try {
                    Get-WindowsCapability -Online -Name "Rsat.BitLocker.Recovery.Tools*" | Add-WindowsCapability -Online -ErrorAction Stop
                    Write-Host "BitLocker RSAT-Tools erfolgreich installiert." -ForegroundColor Green
                } catch {
                     throw "Fehler bei der Installation der BitLocker RSAT-Tools. Bitte manuell installieren und das Skript neu starten."
                }
             } else {
                throw "Die Installation der BitLocker-Tools wurde abgelehnt. Das Skript kann nicht fortgesetzt werden."
             }
        } else {
             Write-Verbose "BitLocker-Tools (repair-bde) sind bereits vorhanden."
        }
        Write-Host "Prüfung der Voraussetzungen abgeschlossen." -ForegroundColor Green
        Write-Host "----------------------------------------------------------------"
    }
    Ensure-Prerequisites
    # --- ENDE: Funktion ---
}

process {
    $vhdMounted = $false; $vhdxPath = $null;

    try {
        Clear-Host
        Write-Host "================================================================" -ForegroundColor Yellow
        Write-Host "      Intelligentes BitLocker Recovery Skript v6.2 (Lokale Rettung)" -ForegroundColor Yellow
        Write-Host "================================================================" -ForegroundColor Yellow
        
        # --- Parameter interaktiv abfragen ---
        $sourceVolume = $null
        while (-not $sourceVolume) {
            if (-not $SourceDrive) { $SourceDrive = Read-Host "Bitte geben Sie den Buchstaben des Quelllaufwerks an (z.B. E)" }
            $sourceVolume = Get-Volume -DriveLetter $SourceDrive.Trim(":") -ErrorAction SilentlyContinue
            if (-not $sourceVolume) { Write-Warning "Laufwerk '$SourceDrive' konnte nicht gefunden werden."; $SourceDrive = $null }
        }
        while (-not ($RecoveryKey -and ($RecoveryKey.Replace("-", "") -match '^\d{48}$'))) {
            $RecoveryKey = Read-Host "Bitte geben Sie den 48-stelligen BitLocker-Wiederherstellungsschlüssel ein"
        }
        $RecoveryKey = $RecoveryKey.Replace("-", "")
        
        while (-not (Test-Path -Path $DestinationFolderPath -PathType Container)) {
            $DestinationFolderPath = Read-Host "Bitte geben Sie einen gültigen, LOKALEN Zielordner an (z.B. F:\Rettung auf Ihrer USB-Festplatte)"
        }
        
        if ([string]::IsNullOrWhiteSpace($VhdxFileName)) { $VhdxFileName = Read-Host "Bitte geben Sie einen Dateinamen für die VHDX-Datei an (ohne .vhdx)" }
        $VhdxFileName = $VhdxFileName.Replace(".vhdx", "")
        $vhdxPath = Join-Path -Path $DestinationFolderPath -ChildPath "$($VhdxFileName).vhdx"
        while (Test-Path -Path $vhdxPath) {
            $choice = Read-Host "Die Zieldatei '$vhdxPath' existiert bereits. [Ü]berschreiben oder [A]bbrechen?"
            switch ($choice.ToUpper()) {
                "Ü" { if ($PSCmdlet.ShouldProcess($vhdxPath, "Alte Datei löschen")) { Remove-Item -Path $vhdxPath -Force; Write-Host "Alte Datei wurde gelöscht." -ForegroundColor Cyan }; break }
                "A" { throw "Aktion vom Benutzer abgebrochen." }
                default { Write-Warning "Ungültige Eingabe."}
            }
        }
        
        $sourcePartition = Get-Partition -DriveLetter $sourceVolume.DriveLetter
        $maxVHDXSize = $sourcePartition.Size
        if ($maxVHDXSize -le 0) { throw "Größe des Quelllaufwerks ist ungültig (0 GB)." }
        $maxVHDXSizeGB = [Math]::Round($maxVHDXSize / 1GB, 2)
        
        $actionSummary = "Datenrettung von '$SourceDrive' in die lokale VHDX '$vhdxPath'."
        Write-Host "`n--- Geplante Aktionen ---"; Write-Host "Zieldatei: $vhdxPath ($maxVHDXSizeGB GB)"; Write-Host "Quelle: $SourceDrive"; Write-Host "----------------------------------------------------------------`n"
        if (-not ($Force) -and -not ($PSCmdlet.ShouldProcess($actionSummary, "Fortfahren?"))) { return }
        
        Write-Verbose "Erstelle VHDX auf dem Pfad: $vhdxPath"
        New-VHD -Path $vhdxPath -SizeBytes $maxVHDXSize -Dynamic
        
        Write-Verbose "Binde VHDX ein..."
        Mount-VHD -Path $vhdxPath; $vhdMounted = $true
        
        $diskNumber = (Get-VHD -Path $vhdxPath).DiskNumber; if (-not $diskNumber) { throw "Konnte die Datenträgernummer der VHDX nicht ermitteln."}
        Initialize-Disk -Number $diskNumber -PartitionStyle GPT
        New-Partition -DiskNumber $diskNumber -UseMaximumSize -AssignDriveLetter
        
        Write-Verbose "Warte, bis das neue Volume bereit ist..."
        $newPartition = Get-Partition -DiskNumber $diskNumber | Where-Object { $_.Type -eq 'Basic' }
        $timer = [System.Diagnostics.Stopwatch]::StartNew()
        while ($timer.Elapsed.TotalSeconds -lt 45) {
            $newVolume = Get-Volume -DriveLetter $newPartition.DriveLetter -ErrorAction SilentlyContinue
            if ($newVolume.HealthStatus -eq 'Healthy') { Write-Verbose "Volume ist bereit."; break }
            Start-Sleep -Seconds 2
        }
        $timer.Stop()
        if ($newVolume.HealthStatus -ne 'Healthy') { throw "Das neue Volume auf der VHDX wurde nach 45 Sekunden nicht ordnungsgemäß initialisiert."}
        
        $destinationRepairDrive = "$($newPartition.DriveLetter):"
        Format-Volume -DriveLetter $newPartition.DriveLetter -FileSystem NTFS -Full -Force -Confirm:$false
        
        Write-Verbose "Gebe Handles auf das Quelllaufwerk '$SourceDrive' frei..."
        $sourceVolume = $null; [GC]::Collect(); Start-Sleep -Seconds 2
        
        Write-Host "`nStarte 'repair-bde'. Dies kann SEHR LANGE dauern." -ForegroundColor Yellow
        repair-bde $SourceDrive $destinationRepairDrive -rp $RecoveryKey -Force
        if ($LASTEXITCODE -ne 0) { throw "'repair-bde' wurde mit Fehlercode '$LASTEXITCODE' beendet." }
        Write-Host "'repair-bde' Prozess erfolgreich abgeschlossen." -ForegroundColor Green
        Write-Host "Ihre geretteten Daten befinden sich nun in der Datei '$vhdxPath'." -ForegroundColor White

    } catch {
        Write-Error "Ein schwerwiegender Fehler ist aufgetreten: $($_.Exception.Message)"
    } finally {
        if ($vhdMounted) { Dismount-VHD -Path $vhdxPath -ErrorAction SilentlyContinue }
        Write-Host "`n================================================================" -ForegroundColor Green
        Write-Host "SKRIPT BEENDET!" -ForegroundColor Green
    }
}
end {
    Write-Verbose "Skriptausführung vollständig beendet."
}
