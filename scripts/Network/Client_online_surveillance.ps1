<#
.SYNOPSIS
Dieses Skript ueberwacht den Status eines Zielhosts durch einen ICMP-Ping und optional durch Port-Knocking.

.DESCRIPTION
Dieses PowerShell-Skript ueberwacht den Status eines Zielhosts, indem es einen ICMP-Ping ausfuehrt und optional Port-Knocking durchfuehrt, um zu ueberpruefen, ob bestimmte Ports erreichbar sind. Es zeigt den aktuellen Status live in einer Tabelle an und speichert die Ergebnisse in einer CSV-Datei.

.PARAMETER Target
Die IP-Adresse oder der Hostname des Zielhosts.

.PARAMETER Port
Der gewuenschte Zielport fuer das Port-Knocking/die Statusmeldung.

.PARAMETER FRQ
Der Zeitabstand in Sekunden, in dem das Ziel ueberprueft wird. Standardwert ist 5 Sekunden.

.PARAMETER TIME
Die Gesamtzeit in Minuten, wie lange das Ziel ueberprueft wird. Standardwert ist 10 Minuten.

.PARAMETER Knock
Gibt an, ob Port-Knocking ausgefuehrt werden soll, unabhaengig vom Erfolg des ICMP-Pings.

.PARAMETER PortsToKnock
Ein Array von Ports, die beim Port-Knocking ueberprueft werden sollen.

.EXAMPLE
.\Client_online_surveillance.ps1 -Target 192.168.1.100 -Port 8080 -FRQ 10 -TIME 300

Ueberwacht den Status des Zielhosts 192.168.1.100 auf Port 8080 alle 10 Sekunden fuer insgesamt 300 Sekunden (5 Minuten).

.EXAMPLE
.\Client_online_surveillance.ps1 -Target google.com -Knock

Ueberwacht den Status des Zielhosts google.com, fuehrt Port-Knocking durch und verwendet die Standardwerte fuer FRQ (5 Sekunden) und TIME (10 Minuten).
#>

param (
    [Parameter(Mandatory = $true, HelpMessage = "Die IP-Adresse oder der Hostname des Zielhosts.")]
    [string]$Target,

    [Parameter(Mandatory = $false)]
    [int]$Port = 0,

    [Parameter(Mandatory = $false)]
    [int]$FRQ = 5,

    [Parameter(Mandatory = $false)]
    [int]$TIME = 10,  

    [Parameter(Mandatory = $false)]
    [switch]$Knock,

    [Parameter(Mandatory = $false)]
    [int[]]$PortsToKnock = @(80, 443, 22, 3389, 445, 139, 53, 161, 389, 636, 3268, 3269, 1433, 1521, 5432, 8080, 8443)
)

function PerformICMPPing {
    param (
        [string]$Target
    )
    try {
        $ping = New-Object System.Net.NetworkInformation.Ping
        # Sende Ping mit einem Timeout von 2000 ms
        $pingResult = $ping.Send($Target, 2000)
        if ($pingResult.Status -eq "Success") {
            Write-Host ("     [OK] ICMP Ping erfolgreich: {0} ms" -f $pingResult.RoundtripTime) -ForegroundColor Green
            return $true
        } else {
            Write-Host "     [ERR] ICMP Ping fehlgeschlagen: $($pingResult.Status)" -ForegroundColor Red
            return $false
        }
    } catch {
        Write-Host "     [ERR] ICMP Ping Fehler: $_" -ForegroundColor Red
        return $false
    }
}

function PerformPortKnocking {
    param (
        [string]$Target,
        [int[]]$PortsToKnock
    )
    
    $knockingTasks = @()
    
    Write-Host "`n[+] Starte Port-Knocking auf $Target..." -ForegroundColor Cyan

    foreach ($p in $PortsToKnock) {
        $knockingTask = {
            param (
                [string]$Target,
                [int]$Port
            )
            try {
                $knockClient = New-Object System.Net.Sockets.TcpClient
                # Verwende ein kurzes Timeout von 1000ms ueber BeginConnect
                $asyncResult = $knockClient.BeginConnect($Target, $Port, $null, $null)
                $wait = $asyncResult.AsyncWaitHandle.WaitOne(1000)
                if ($wait -and $knockClient.Connected) {
                    $knockClient.EndConnect($asyncResult)
                    $knockClient.Close()
                    return $Port
                } else {
                    $knockClient.Close()
                    return $null
                }
            } catch {
                return $null
            }
        }

        $knockingTasks += Start-Job -ScriptBlock $knockingTask -ArgumentList $Target, $p
    }

    # Interaktiver Fortschrittsbalken waehrend des Wartens auf die Jobs
    $totalJobs = $knockingTasks.Count
    $completedJobs = 0
    
    while ($completedJobs -lt $totalJobs) {
        $completedJobs = ($knockingTasks | Where-Object { $_.State -ne "Running" }).Count
        $percent = [int](($completedJobs / $totalJobs) * 100)
        Write-Progress -Activity "Port Knocking Ueberpruefung" -Status "$completedJobs von $totalJobs Ports ueberprueft ($percent%)" -PercentComplete $percent
        Start-Sleep -Milliseconds 100
    }
    # Schliesse Fortschrittsbalken ab
    Write-Progress -Activity "Port Knocking Ueberpruefung" -Completed

    # Ergebnisse einsammeln
    $results = $knockingTasks | Receive-Job | Where-Object { $_ -ne $null }
    
    # Jobs aufraeumen
    $knockingTasks | Remove-Job -Force

    $successPorts = $results | Where-Object { $_ -is [int] }
    $failedPorts = $PortsToKnock | Where-Object { $_ -notin $successPorts }

    Write-Host "[OK] Port-Knocking abgeschlossen.`n" -ForegroundColor Cyan
    
    if ($successPorts.Count -gt 0) {
        Write-Host "Erfolgreich geoeffnete/erreichbare Ports:" -ForegroundColor Green
        foreach ($sp in $successPorts) {
            Write-Host "  - Port $sp (OFFEN)" -ForegroundColor Green
        }
    } else {
        Write-Host "Keine Ports erfolgreich erreicht." -ForegroundColor Yellow
    }

    if ($failedPorts.Count -gt 0) {
        $failedPortsStr = $failedPorts -join ", "
        Write-Host "Geschlossene/nicht erreichbare Ports:" -ForegroundColor DarkGray
        Write-Host "  - Ports: $failedPortsStr" -ForegroundColor DarkGray
    }
    
    return $successPorts
}

function ReportStatustoScreen {
    param (
        [string]$status,
        [int]$Port
    )
    if ($Port -le 0) { return }
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        # Verbindung mit Timeout aufbauen
        $asyncResult = $client.BeginConnect($Target, $Port, $null, $null)
        $wait = $asyncResult.AsyncWaitHandle.WaitOne(1000)
        if ($wait -and $client.Connected) {
            $client.EndConnect($asyncResult)
            $clientAddress = $client.Client.RemoteEndPoint.Address
            
            $stream = $client.GetStream()
            $writer = New-Object System.IO.StreamWriter $stream
            $writer.WriteLine($status)
            $writer.Flush()
            $client.Close()
            
            Write-Host "     -> Status erfolgreich an ${clientAddress}:$Port gemeldet" -ForegroundColor DarkGray
        } else {
            $client.Close()
        }
    } catch {
        # Fehler beim optionalen Port-Melder leise uebergehen, um Dauerschleife nicht zu stoehren
    }
}

function ReportStatustoFile {
    param (
        [string]$status,
        [int]$Port
    )
    $timeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $tableEntry = [PSCustomObject]@{
        "Timestamp" = $timeStamp
        "Status"    = $status
        "Port"      = $Port
    }
    
    # Premium-konsolenbasiertes Logging mit Farben
    $color = if ($status -eq "Online") { "Green" } else { "Red" }
    Write-Host "[$timeStamp] " -NoNewline -ForegroundColor Gray
    Write-Host "Status: " -NoNewline
    Write-Host "$status" -ForegroundColor $color -NoNewline
    Write-Host " | Zielhost: $Target | Port: $Port"

    # In CSV-Datei schreiben
    try {
        $tableEntry | Export-Csv -Append -NoTypeInformation -Path "StatusReport.csv" -Encoding UTF8
    } catch {
        Write-Warning "Fehler beim Schreiben in StatusReport.csv: $_"
    }
}

function Main {
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host "        CLIENT ONLINE SURVEILLANCE & KNOCKER      " -ForegroundColor Cyan
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host "Zielhost: $Target" -ForegroundColor Yellow
    Write-Host "Startzeit: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n" -ForegroundColor Gray

    # Ersten ICMP-Ping ausfuehren
    Write-Host "[+] Fuehre initialen Ping durch..." -ForegroundColor Cyan
    $pingSuccess = PerformICMPPing -Target $Target
    $status = if ($pingSuccess) { "Online" } else { "Offline" }
    ReportStatustoFile -status $status -Port $Port
    
    # Wenn Ping fehlschlaegt oder -Knock erzwungen wird, fuehre Port-Knocking durch
    if (-not $pingSuccess -or $Knock) {
        $successPorts = PerformPortKnocking -Target $Target -PortsToKnock $PortsToKnock
        
        # Wenn kein Zielport angegeben wurde, interaktive Auswahl anbieten
        if ($Port -eq 0) {
            if ($successPorts.Count -gt 0) {
                Write-Host "`nVerfuegbare offene Ports: $($successPorts -join ', ')" -ForegroundColor Yellow
                $userInput = Read-Host "Waehlen Sie einen Zielport fuer die Ueberwachung (Standard: $($successPorts[0]))"
                if ([string]::IsNullOrWhiteSpace($userInput)) {
                    $Port = $successPorts[0]
                } else {
                    if ($userInput -match '^\d+$') {
                        $Port = [int]$userInput
                    } else {
                        $Port = $successPorts[0]
                    }
                }
            } else {
                $userInput = Read-Host "Keine offenen Ports gefunden. Geben Sie einen Zielport manuell ein oder druecken Sie Enter fuer keinen"
                if (-not [string]::IsNullOrWhiteSpace($userInput) -and ($userInput -match '^\d+$')) {
                    $Port = [int]$userInput
                } else {
                    $Port = 0
                }
            }
        } else {
            Write-Host "`nZielport ist auf $Port konfiguriert." -ForegroundColor Cyan
        }
    }

    # Kontinuierliche Ueberwachungsschleife starten
    Write-Host "`n[+] Starte Ueberwachung von $Target alle $FRQ Sekunden fuer insgesamt $TIME Minuten..." -ForegroundColor Cyan
    Write-Host "    (Ergebnisse werden zusaetzlich in 'StatusReport.csv' gespeichert)`n" -ForegroundColor DarkGray
    
    $endTime = (Get-Date).AddMinutes($TIME)
    
    while ((Get-Date) -le $endTime) {
        $pingResult = PerformICMPPing -Target $Target
        $loopStatus = if ($pingResult) { "Online" } else { "Offline" }
        ReportStatustoFile -status $loopStatus -Port $Port
        
        # Optional: Status an den Zielport uebertragen
        if ($Port -gt 0) {
            ReportStatustoScreen -status $loopStatus -Port $Port
        }
        
        Start-Sleep -Seconds $FRQ
    }
    
    Write-Host "`n[OK] Ueberwachung beendet um $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')." -ForegroundColor Green
    Write-Host "==================================================" -ForegroundColor Cyan
}

# Starte das Hauptprogramm
Main
