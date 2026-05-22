<#
.SYNOPSIS
Ueberprueft den freien Festplattenspeicherplatz auf lokalen und Netzwerk-Computern.

.DESCRIPTION
Dieses PowerShell-Skript analysiert den verfuegbaren Speicherplatz auf angegebenen Laufwerken von einer Liste von Computern.
Es minimiert die Netzwerkbelastung, indem es alle Laufwerkinformationen in einem einzigen CIM-Aufruf pro Computer abfragt.
Es bietet eine farblich hervorgehobene Konsolenausgabe und eine optionale CSV-Protokollierung.

.PARAMETER ComputerName
Ein Array von Computernamen oder IP-Adressen. Standardmaessig wird versucht, die Hostnamen aus 'Computers.txt' zu laden. Falls dies nicht existiert, wird 'localhost' verwendet.

.PARAMETER DriveLetters
Die zu pruefenden Laufwerksbuchstaben (z. B. 'C', 'D', 'E'). Standard ist C, D, E, F, G.

.PARAMETER WarningThresholdPercent
Der Prozentsatz an freiem Speicherplatz, ab dem eine Warnung (Gelb) ausgegeben wird. Standard ist 15%.

.PARAMETER CriticalThresholdPercent
Der Prozentsatz an freiem Speicherplatz, ab dem ein kritischer Status (Rot) ausgegeben wird. Standard ist 5%.

.PARAMETER ExportPath
Optionaler Pfad fuer den Export der Ergebnisse in eine CSV-Datei.

.EXAMPLE
.\Get-HDDSpaceReport.ps1 -ComputerName "Server01", "localhost" -DriveLetters C, D
#>

param (
    [Parameter(Mandatory = $false, Position = 0)]
    [string[]]$ComputerName,

    [Parameter(Mandatory = $false)]
    [string[]]$DriveLetters = @('C', 'D', 'E', 'F', 'G'),

    [Parameter(Mandatory = $false)]
    [int]$WarningThresholdPercent = 15,

    [Parameter(Mandatory = $false)]
    [int]$CriticalThresholdPercent = 5,

    [Parameter(Mandatory = $false)]
    [string]$ExportPath
)

# Standardwert fuer ComputerName, falls nicht angegeben
if (-not $ComputerName) {
    $computersFile = Join-Path $PSScriptRoot "Computers.txt"
    if (Test-Path $computersFile) {
        $ComputerName = Get-Content $computersFile | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and $_ -notmatch '^#' }
    }
    
    if (-not $ComputerName) {
        $ComputerName = @("localhost")
    }
}

# Normalisiere Laufwerksbuchstaben zu 'C:', 'D:' etc.
$normalizedDrives = $DriveLetters | ForEach-Object { "$($_[0].ToString().ToUpper()):" }

Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "         HDD SPACE CHECKER & MONITORING           " -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "Ziel-Computer: $($ComputerName -join ', ')" -ForegroundColor Yellow
Write-Host "Startzeit:     $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n" -ForegroundColor Gray

$reportData = @()

foreach ($computer in $ComputerName) {
    $computer = $computer.Trim()
    Write-Host "[+] Pruefe Host: $computer..." -ForegroundColor Cyan

    # Ping-Test vor der Abfrage, um Timeout-Verzoegerungen zu verhindern
    $pingSuccess = $false
    if ($computer -eq "localhost" -or $computer -eq "127.0.0.1" -or $computer -eq $env:COMPUTERNAME) {
        $pingSuccess = $true
    } else {
        try {
            $ping = New-Object System.Net.NetworkInformation.Ping
            $reply = $ping.Send($computer, 1000)
            if ($reply.Status -eq "Success") {
                $pingSuccess = $true
            }
        } catch {
            $pingSuccess = $false
        }
    }

    if (-not $pingSuccess) {
        Write-Host "  [ERR] Computer ist offline oder nicht erreichbar (Ping fehlgeschlagen)." -ForegroundColor Red
        continue
    }

    try {
        # Verwende Get-CimInstance (moderner Standard) anstelle von Get-WmiObject (obsolet)
        # Alle Laufwerke in einem einzigen Aufruf abfragen, um Netzwerk-Roundtrips zu minimieren!
        $disks = $null
        if ($computer -eq "localhost" -or $computer -eq "127.0.0.1" -or $computer -eq $env:COMPUTERNAME) {
            # Lokale Abfrage erfolgt ueber die extrem schnelle und sichere .NET DriveInfo-API.
            # Vermeidet WMI-Haenger durch unbereite oder getrennte Netzfreigaben komplett!
            $disks = [System.IO.DriveInfo]::GetDrives() | Where-Object { 
                $_.DriveType -eq [System.IO.DriveType]::Fixed -and $_.IsReady -and $normalizedDrives -contains $_.Name.Substring(0, 2)
            } | ForEach-Object {
                [PSCustomObject]@{
                    DeviceID  = $_.Name.Substring(0, 2)
                    Size      = $_.TotalSize
                    FreeSpace = $_.TotalFreeSpace
                }
            }
        } else {
            # Remote-Abfrage: Versuche DCOM (abwaertskompatibel wie Get-WmiObject)
            try {
                $sessionOption = New-CimSessionOption -Protocol Dcom
                $session = New-CimSession -ComputerName $computer -SessionOption $sessionOption -ErrorAction Stop
                $disks = Get-CimInstance -ClassName Win32_LogicalDisk -CimSession $session -ErrorAction Stop | 
                         Where-Object { $normalizedDrives -contains $_.DeviceID }
                Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue
            } catch {
                # Fallback auf standardmaessiges WSMan/WinRM
                $disks = Get-CimInstance -ClassName Win32_LogicalDisk -ComputerName $computer -ErrorAction Stop | 
                         Where-Object { $normalizedDrives -contains $_.DeviceID }
            }
        }

        if (-not $disks) {
            Write-Host "  [WARN] Keine der angegebenen Laufwerke ($($DriveLetters -join ', ')) auf dem Host gefunden." -ForegroundColor Yellow
            continue
        }

        foreach ($disk in $disks) {
            $sizeGB = [math]::Round($disk.Size / 1GB, 2)
            $freeGB = [math]::Round($disk.FreeSpace / 1GB, 2)
            
            if ($disk.Size -gt 0) {
                $freePercent = [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 1)
            } else {
                $freePercent = 0
            }

            # Bestimme Statusfarbe
            $status = "OK"
            $color = "Green"
            if ($freePercent -le $CriticalThresholdPercent) {
                $status = "CRITICAL"
                $color = "Red"
            } elseif ($freePercent -le $WarningThresholdPercent) {
                $status = "WARNING"
                $color = "Yellow"
            }

            Write-Host "  - Laufwerk $($disk.DeviceID) " -NoNewline
            Write-Host "Status: $status" -ForegroundColor $color -NoNewline
            Write-Host " | Gesamt: $sizeGB GB | Frei: $freeGB GB ($freePercent%)"

            $reportData += [PSCustomObject]@{
                "Computer"    = $computer
                "Drive"       = $disk.DeviceID
                "TotalSizeGB" = $sizeGB
                "FreeSpaceGB" = $freeGB
                "FreePercent" = $freePercent
                "Status"      = $status
                "Timestamp"   = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            }
        }
    } catch {
        Write-Host "  [ERR] Fehler beim Abrufen der Festplatten-Daten: $_" -ForegroundColor Red
    }
}

# Export der Daten, falls ein Pfad angegeben wurde
if ($ExportPath -and $reportData.Count -gt 0) {
    try {
        $reportData | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
        Write-Host "`n[OK] Bericht erfolgreich exportiert nach: $ExportPath" -ForegroundColor Green
    } catch {
        Write-Warning "Fehler beim Exportieren des Berichts nach ${ExportPath}: $_"
    }
}

Write-Host "`n==================================================" -ForegroundColor Cyan
Write-Host "Ueberpruefung abgeschlossen." -ForegroundColor Green
Write-Host "==================================================" -ForegroundColor Cyan
