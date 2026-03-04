<#
.SYNOPSIS
    New-NetworkTopologyReport.ps1
    Erstellt einen HTML-Report der lokalen Netzwerk-Topologie mit optionalen,
    differenzierten Port-Scans UND einem interaktiven Modus für
    gezielte Einzel-Scans (inkl. SMTP-Ports).
    (PS-Coding Assistent)

.DESCRIPTION
    Kombiniert passive und (optional) aktive Methoden, um einen detaillierten
    Report der lokalen Netzwerkumgebung zu erstellen.

    MODI:
    1. Passiv (Standard): Erstellt einen HTML-Report via ARP-Analyse (Hop 0)
       und TraceRoute (Hop 1-N).
    2. Aktiv (Optional): Schalter -ScanInfrastructure und/oder
       -ScanLocalClientsAndServers führen Port-Scans für die Report-Erstellung
       durch.
    3. Interaktiv (NEU): Der Schalter -InteractiveScan startet NACH dem
       passiven Scan ein Menü zur gezielten Auswahl und Prüfung einzelner
       Geräte und Dienste.

.PARAMETER HtmlReportPath
    Der Speicherort für den finalen HTML-Report.
    Standard: ".\NetworkTopologyReport.html"

.PARAMETER TraceTarget
    Das Ziel für die TraceRoute-Analyse. Standard: "8.8.8.8"

.PARAMETER MaxHops
    Maximale Tiefe der TraceRoute-Analyse. Standard: 8

.PARAMETER ScanInfrastructure
    [Switch] Aktiviert Port-Scan für *Infrastruktur-Geräte* (Report-Erstellung).

.PARAMETER InfrastructurePorts
    Ports für -ScanInfrastructure.
    Standard: @(22, 23, 80, 443, 161) (SSH, Telnet, HTTP, HTTPS, SNMP)

.PARAMETER ScanLocalClientsAndServers
    [Switch] ACHTUNG: Aktiviert Port-Scan für *Clients/Server* in Hop 0
    (Report-Erstellung). KANN ALARME AUSLÖSEN.

.PARAMETER ClientServerPorts
    Ports für -ScanLocalClientsAndServers.
    Standard: @(21, 80, 135, 139, 443, 445, 3389) (FTP, HTTP, MSRPC, NetBIOS,
              HTTPS, SMB, RDP)

.PARAMETER InteractiveScan
    [Switch] Startet nach dem passiven Scan ein interaktives Menü,
    um einzelne Geräte gezielt auf ausgewählte Ports zu scannen.
    Ignoriert die -Scan... Parameter und führt den Scan interaktiv durch.

.EXAMPLE
    PS C:\> .\New-NetworkTopologyReport.ps1
    Erstellt "NetworkTopologyReport.html" (rein passiv).

.EXAMPLE
    PS C:\> .\New-NetworkTopologyReport.ps1 -ScanInfrastructure
    Erstellt Report inkl. Port-Scan für Switche/Router.

.EXAMPLE
    PS C:\> .\New-NetworkTopologyReport.ps1 -InteractiveScan
    Erstellt den passiven Report UND startet danach das Scan-Menü,
    um z.B. SMTP-Server (Port 25/587) gezielt zu suchen.
#>

[CmdletBinding()]
param (
    [string]$HtmlReportPath = ".\NetworkTopologyReport.html",
    [string]$TraceTarget = "8.8.8.8",
    [int]$MaxHops = 8,

    [switch]$ScanInfrastructure,
    [int[]]$InfrastructurePorts = @(22, 23, 80, 443, 161),

    [switch]$ScanLocalClientsAndServers,
    [int[]]$ClientServerPorts = @(21, 80, 135, 139, 443, 445, 3389),
    
    [switch]$InteractiveScan
)

#--------------------------------------------------------------------------
# 1. Definitionen und Vorbereitungen
#--------------------------------------------------------------------------
Write-Verbose "Skript gestartet. Modus: HTML-Report."
if ($InteractiveScan) {
    Write-Host "Interaktiver Modus ist aktiviert." -ForegroundColor Cyan
    Write-Host "Die Parameter -ScanInfrastructure und -ScanLocalClientsAndServers werden ignoriert."
    $ScanInfrastructure = $false
    $ScanLocalClientsAndServers = $false
} elseif ($ScanLocalClientsAndServers) {
    Write-Warning "!!! SICHERHEITSWARNUNG !!!"
    Write-Warning "ScanLocalClientsAndServers ist aktiviert. Dies kann IDS/IPS-Alarme auslösen."
    Start-Sleep -Seconds 2
}

# OUI-Liste (Hersteller)
$OUI_List = @{
    # Cisco (Catalyst)
    "00-00-0C" = "Cisco"; "00-10-7B" = "Cisco"; "00-05-9A" = "Cisco";
    "00-07-0D" = "Cisco"; "00-E0-B0" = "Cisco"; "00-1A-E3" = "Cisco";
    "00-26-98" = "Cisco"; "50-3D-E5" = "Cisco"; "B4-A4-E3" = "Cisco";
    "C8-F9-F9" = "Cisco";
    # Linksys
    "00-14-BF" = "Linksys"; "00-26-F2" = "Linksys"; "00-0F-66" = "Linksys";
    # Aruba (HPE)
    "00-0B-86" = "Aruba (HPE)"; "00-1A-1E" = "Aruba (HPE)"; "E8-04-10" = "Aruba (HPE)";
    # 3Com (HPE)
    "00-01-80" = "3Com"; "00-04-75" = "3Com"; "00-50-DA" = "3Com";
    # Microsens
    "00-0A-D3" = "MICROSENS GmbH"; "00-10-E7" = "MICROSENS GmbH";
    # Andere
    "08-00-09" = "HPE (ProCurve)"; "00-60-B9" = "HPE (ProCurve)";
    "00-01-E8" = "Dell"; "00-12-3F" = "Dell";
    "00-05-85" = "Juniper"; "00-90-69" = "Brocade";
    "00-E0-52" = "Extreme"; "00-04-96" = "Extreme (Enterasys)";
    # Clients / Drucker / VM
    "00-05-69" = "VMware"; "00-50-56" = "VMware"; "00-0C-29" = "VMware";
    "00-1C-42" = "Microsoft (Hyper-V)"; "00-03-FF" = "Microsoft (Hyper-V)";
    "00-80-77" = "HP (Printer)"; "00-00-E2" = "HP (Printer)";
    "00-E0-4C" = "Realtek (Client NIC)"; "00-1B-21" = "Intel (Client NIC)";
}

# Definition der Dienste für den interaktiven Modus (INKLUSIVE SMTP)
$InteractiveServiceList = @(
    # Management & Web
    [PSCustomObject]@{ Name = "SSH"; Port = 22; Zweck = "Sicheres Management (Infrastruktur)" },
    [PSCustomObject]@{ Name = "Telnet"; Port = 23; Zweck = "Unsicheres Management (Legacy)" },
    [PSCustomObject]@{ Name = "HTTP"; Port = 80; Zweck = "Web-Oberfläche (Standard)" },
    [PSCustomObject]@{ Name = "HTTPS"; Port = 443; Zweck = "Sichere Web-Oberfläche" },
    [PSCustomObject]@{ Name = "SNMP"; Port = 161; Zweck = "Netzwerk-Monitoring" },
    
    # Windows-Dienste
    [PSCustomObject]@{ Name = "RDP"; Port = 3389; Zweck = "Remote Desktop (Windows)" },
    [PSCustomObject]@{ Name = "SMB"; Port = 445; Zweck = "Windows-Dateifreigabe" },
    [PSCustomObject]@{ Name = "NetBIOS-SSN"; Port = 139; Zweck = "Legacy-Dateifreigabe (NetBIOS)" },
    [PSCustomObject]@{ Name = "MSRPC"; Port = 135; Zweck = "Windows RPC (Dienste-Endpunkt)" },
    [PSCustomObject]@{ Name = "FTP"; Port = 21; Zweck = "Datei-Transfer (Unsicher)" },

    # E-Mail Dienste (AKTUALISIERT)
    [PSCustomObject]@{ Name = "SMTP"; Port = 25; Zweck = "E-Mail-Übertragung (Server-zu-Server)" },
    [PSCustomObject]@{ Name = "SMTPS"; Port = 465; Zweck = "E-Mail-Übertragung (Sicher/Legacy)" },
    [PSCustomObject]@{ Name = "SMTP-SUBMIT"; Port = 587; Zweck = "E-Mail-Versand (Client-zu-Server)" }
)

# Liste für alle finalen Objekte, die in den Report kommen
$reportObjects = [System.Collections.Generic.List[object]]::new()
$allDiscoveredIPs = [System.Collections.Generic.List[string]]::new()

#--------------------------------------------------------------------------
# 2. Hilfsfunktionen
#--------------------------------------------------------------------------

function Format-MacAddress ($Mac) {
    if ($Mac -match '([0-9A-Fa-f]{2})' -and $Mac.Length -ge 12) {
        return ($Mac -replace '([0-9A-Fa-f]{2})(?!$)', '$1-').Substring(0, 17).ToUpper()
    }
    return $Mac.ToUpper()
}

function Get-MacPrefix ($FormattedMac) {
    if ($FormattedMac -and $FormattedMac.Length -ge 8) {
        return $FormattedMac.Substring(0, 8)
    }
    return $null
}

# Zentrale Funktion: Ermittelt Details zu EINER IP-Adresse
function Get-DeviceInfo ($IPAddress, [int[]]$PortsToScanList = $null) {
    $mac = "N/A"
    $prefix = "N/A"
    $manufacturer = "Unbekannt"
    $hostname = "N/A"

    # 1. MAC & Hersteller (funktioniert nur für lokales VLAN / Hop 0)
    try {
        $neighbor = Get-NetNeighbor -IPAddress $IPAddress -ErrorAction SilentlyContinue
        if ($neighbor -and $neighbor.LinkLayerAddress) {
            $mac = Format-MacAddress $neighbor.LinkLayerAddress
            $prefix = Get-MacPrefix $mac
            if ($prefix -and $OUI_List.ContainsKey($prefix)) {
                $manufacturer = $OUI_List[$prefix]
            }
        }
    } catch { }

    # 2. Hostname (Reverse DNS)
    try {
        $dnsResult = Resolve-DnsName -Name $IPAddress -Type PTR -DnsOnly -ErrorAction SilentlyContinue
        if ($dnsResult -and $dnsResult.NameHost) {
            $hostname = $dnsResult.NameHost
        }
    } catch {
        $hostname = "Fehler bei rDNS"
    }

    # 3. Gezielter Port-Scan (für Report-Modus)
    $openPorts = "N/A"
    if ($InteractiveScan) {
        $openPorts = "N/A (Siehe Interaktiver Modus)"
    } elseif ($PortsToScanList) {
        $openPortsList = [System.Collections.Generic.List[string]]::new()
        $portListString = [string]::Join(', ', $PortsToScanList)
        Write-Verbose "  -> Starte Report-Port-Scan für $IPAddress (Ports: $portListString)"
        
        foreach ($port in $PortsToScanList) {
            $tnc = Test-NetConnection -ComputerName $IPAddress -Port $port -WarningAction SilentlyContinue -InformationLevel Quiet
            if ($tnc.TcpTestSucceeded) {
                $openPortsList.Add($port.ToString())
            }
        }
        $openPorts = if ($openPortsList.Count -gt 0) {
            [string]::Join(', ', $openPortsList)
        } else {
            "Keine (von $portListString)"
        }
    }

    # 4. Objekt zurückgeben
    return [PSCustomObject]@{
        Hop          = "N/A"
        IP_Adresse   = $IPAddress
        Hostname     = $hostname
        MAC_Adresse  = $mac
        Hersteller   = $manufacturer
        Offene_Ports = $openPorts
    }
}

# Funktion für den interaktiven Modus
function Start-InteractiveScanMenu ($DeviceList) {
    
    while ($true) {
        Clear-Host
        Write-Host "--- Interaktiver Scan: Geräteauswahl ---" -ForegroundColor Yellow
        Write-Host "Die folgenden Geräte wurden im passiven Scan gefunden:"
        Write-Host ""
        
        # Menü 1: Geräteauswahl
        $devices = $DeviceList | Where-Object { $_.Hop -ne "Host" }
        for ($i = 0; $i -lt $devices.Count; $i++) {
            $device = $devices[$i]
            Write-Host (" [{0,2}] {1,-15} ({2}) - [{3}] (Hop: {4})" -f $i, $device.IP_Adresse, $device.Hostname, $device.Hersteller, $device.Hop)
        }
        Write-Host ""
        $input = Read-Host "Nummer des zu scannenden Geräts eingeben (oder 'q' zum Beenden)"
        
        if ($input -eq 'q') {
            break
        }
        
        $selectedDevice = $null
        if ($input -match '^\d+$' -and [int]$input -ge 0 -and [int]$input -lt $devices.Count) {
            $selectedDevice = $devices[[int]$input]
        } else {
            Write-Host "Ungültige Auswahl. Bitte eine Nummer aus der Liste eingeben." -ForegroundColor Red
            Start-Sleep -Seconds 2
            continue
        }
        
        # Menü 2: Dienstauswahl (via Out-GridView)
        Write-Host ""
        Write-Host "Öffne Dienstauswahl-Fenster für $($selectedDevice.IP_Adresse)..." -ForegroundColor Cyan
        try {
            $servicesToScan = $Global:InteractiveServiceList | 
                              Out-GridView -Title "Wähle Dienste für $($selectedDevice.IP_Adresse)" -OutputMode Multiple
        } catch {
            Write-Warning "Out-GridView konnte nicht gestartet werden (evtl. in Core-Sitzung?)."
            Write-Warning "Breche interaktiven Scan ab."
            break
        }
        
        if (-not $servicesToScan) {
            Write-Host "Keine Dienste ausgewählt. Kehre zum Hauptmenü zurück." -ForegroundColor Yellow
            Start-Sleep -Seconds 2
            continue
        }
        
        # Durchführung & Ergebnis
        Clear-Host
        Write-Host "--- Scan-Ergebnis für: $($selectedDevice.IP_Adresse) ($($selectedDevice.Hostname)) ---" -ForegroundColor Yellow
        Write-Host ""
        
        foreach ($service in $servicesToScan) {
            Write-Host (" -> Prüfe Port {0,-5} ({1,-12})... " -f $service.Port, $service.Name) -NoNewline
            
            $tnc = Test-NetConnection -ComputerName $selectedDevice.IP_Adresse -Port $service.Port -WarningAction SilentlyContinue -InformationLevel Quiet
            
            if ($tnc.TcpTestSucceeded) {
                Write-Host "[ OFFEN ]" -ForegroundColor Green
                Write-Host ("    Bewertung/Zweck: $($service.Zweck)") -ForegroundColor Gray
            } else {
                Write-Host "[ GESCHLOSSEN ]" -ForegroundColor Red
            }
        }
        
        Write-Host ""
        Read-Host "Enter drücken, um zum Hauptmenü zurückzukehren..."
    }
    
    Write-Host "Interaktiver Modus beendet."
}


#--------------------------------------------------------------------------
# 3. Lokalen Host ermitteln
#--------------------------------------------------------------------------
Write-Verbose "Ermittle lokale IP-Konfiguration..."
$ipConfig = Get-NetIPConfiguration | Where-Object { $_.IPv4DefaultGateway } | Select-Object -First 1
if (-not $ipConfig) {
    Write-Error "Konnte keine aktive Netzwerkverbindung mit Standard-Gateway finden."
    return
}

$localIP = ($ipConfig.IPv4Address | Select-Object -First 1).IPAddress
$reportObjects.Add([PSCustomObject]@{
        Hop          = "Host"
        IP_Adresse   = $localIP
        Hostname     = $env:COMPUTERNAME
        MAC_Adresse  = ($ipConfig.Interface | Get-NetAdapter).MacAddress
        Hersteller   = "Lokaler PC"
        Offene_Ports = "N/A"
    })
$allDiscoveredIPs.Add($localIP)

#--------------------------------------------------------------------------
# 4. Hop 0 (Lokales VLAN / "Clients im Umkreis")
#--------------------------------------------------------------------------
Write-Verbose "Lese ARP-Tabelle (Get-NetNeighbor) für Hop 0..."
try {
    $arpIPs = Get-NetNeighbor -AddressFamily IPv4 | 
              Where-Object { $_.State -ne "Incomplete" -and $_.LinkLayerAddress } | 
              Select-Object -ExpandProperty IPAddress -Unique

    foreach ($ip in $arpIPs) {
        if ($ip -notin $allDiscoveredIPs) {
            Write-Verbose "Analysiere Hop 0 Gerät: $ip"
            
            # Basis-Infos holen
            $deviceInfo = Get-DeviceInfo $ip
            
            # Port-Liste für Report-Modus bestimmen
            $portsForThisDevice = $null
            $isInfrastructure = $deviceInfo.Hersteller -ne "Unbekannt"
            
            if ($ScanInfrastructure -and $isInfrastructure) {
                $portsForThisDevice = $InfrastructurePorts
            } elseif ($ScanLocalClientsAndServers -and (-not $isInfrastructure)) {
                $portsForThisDevice = $ClientServerPorts
            }
            
            # Finale Infos holen (ggf. mit Port-Scan für Report)
            $finalDeviceInfo = Get-DeviceInfo $ip $portsForThisDevice
            $finalDeviceInfo.Hop = "0 (Lokales VLAN)"
            
            $reportObjects.Add($finalDeviceInfo)
            $allDiscoveredIPs.Add($ip)
        }
    }
} catch {
    Write-Warning "Fehler beim Lesen der ARP-Tabelle (Hop 0): $_"
}

#--------------------------------------------------------------------------
# 5. Hop 1-N (Datenpfad)
#--------------------------------------------------------------------------
Write-Verbose "Führe gezielte TraceRoute zu '$TraceTarget' mit max. $MaxHops Hops aus..."
try {
    $trace = Test-NetConnection -ComputerName $TraceTarget -TraceRoute -Hops $MaxHops -WarningAction SilentlyContinue
    
    if ($trace -and $trace.TraceRoute) {
        for ($i = 0; $i -lt $trace.TraceRoute.Count; $i++) {
            $hopNumber = $i + 1
            $hopIP = $trace.TraceRoute[$i].IPAddress
            
            if ($hopIP -ne "0.0.0.0" -and $hopIP -notlike "*.*.*.* *" -and $hopIP -notin $allDiscoveredIPs) {
                Write-Verbose "Analysiere Hop $hopNumber Gerät: $hopIP"
                
                $portsForThisDevice = $null
                if ($ScanInfrastructure) {
                    $portsForThisDevice = $InfrastructurePorts
                }
                
                $deviceInfo = Get-DeviceInfo $ip $portsForThisDevice
                $deviceInfo.Hop = $hopNumber.ToString()
                
                $reportObjects.Add($deviceInfo)
                $allDiscoveredIPs.Add($hopIP)
            }
        }
    } else {
        Write-Warning "Test-NetConnection -TraceRoute hat keine Ergebnisse geliefert."
    }
} catch {
    Write-Error "Schwerer Fehler bei Test-NetConnection: $_"
}

#--------------------------------------------------------------------------
# 6. HTML-Report erstellen
#--------------------------------------------------------------------------
Write-Verbose "Erstelle HTML-Report unter: $HtmlReportPath"

# CSS für den Report
$css = @"
<style>
    body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background-color: #fdfdfd; }
    h1 { color: #2E4053; }
    h2 { color: #566573; border-bottom: 2px solid #AEB6BF; padding-bottom: 5px; }
    table { border-collapse: collapse; width: 100%; margin-top: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
    th, td { border: 1px solid #ddd; padding: 10px 12px; text-align: left; }
    th { background-color: #566573; color: white; }
    tr:nth-child(even) { background-color: #f8f9f9; }
    tr:hover { background-color: #eaf2f8; }
    .header { margin-bottom: 20px; padding: 10px; background-color: #f4f6f6; border-radius: 5px; }
    .footer { margin-top: 30px; font-size: 0.8em; color: #888; }
    .warn { color: #C0392B; font-weight: bold; }
</style>
"@

# Daten aufbereiten und sortieren
$sortedData = $reportObjects | Sort-Object { [int](($_.Hop -split ' ')[0] -replace 'Host', '-1') }, IP_Adresse

# Scan-Status für den Header
$scanStatus = "Passiv (Kein Port-Scan)"
if ($ScanInfrastructure -and $ScanLocalClientsAndServers) {
    $scanStatus = "Vollständig Aktiv (Infrastruktur + Clients)"
} elseif ($ScanInfrastructure) {
    $scanStatus = "Aktiv (Nur Infrastruktur)"
} elseif ($ScanLocalClientsAndServers) {
    $scanStatus = "<span class='warn'>Aktiv (Nur Clients/Server)</span>"
} elseif ($InteractiveScan) {
    $scanStatus = "Passiv (Interaktives Menü folgt)"
}

# HTML erstellen
$htmlBody = $sortedData | ConvertTo-Html -Fragment
$title = "Netzwerk-Topologie Report ($($env:COMPUTERNAME))"

$htmlOutput = ConvertTo-Html -Head ($css + "<title>$title</title>") -Body (
    "<div class='header'><h1>$title</h1>" +
    "Erstellt am: $(Get-Date)<br>" +
    "Trace-Ziel: $TraceTarget<br>" +
    "<b>Scan-Modus: $scanStatus</b>" +
    "</div>" +
    "<h2>Topologie-Übersicht</h2>" +
    $htmlBody +
    "<div class='footer'>Report generiert von New-NetworkTopologyReport.ps1</div>"
)

try {
    $htmlOutput | Out-File -FilePath $HtmlReportPath -Encoding UTF8
    Write-Output "---"
    Write-Output "✅ HTML-Report erfolgreich erstellt:"
    Write-Output (Resolve-Path $HtmlReportPath).Path
    
} catch {
    Write-Error "Fehler beim Speichern des HTML-Reports: $_"
}

#--------------------------------------------------------------------------
# 7. Interaktiven Modus starten
#--------------------------------------------------------------------------
if ($InteractiveScan) {
    Write-Output ""
    Write-Output "Starte interaktiven Scan-Modus..." -ForegroundColor Cyan
    Start-Sleep -Seconds 1
    # Rufe die neue Funktion mit den sortierten Daten auf
    Start-InteractiveScanMenu -DeviceList $sortedData
}

Write-Verbose "Skript beendet."
