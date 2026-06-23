<#
.SYNOPSIS
    Invoke-WOLManager.ps1 - Umfassendes Wake-on-LAN Management Script.
    
.DESCRIPTION
    Dieses Skript sendet Wake-on-LAN (WOL) Magic Packets an Computer.
    Es unterstützt:
    1. Intelligente Eingabe: Erkennt automatisch ob Name oder MAC eingegeben wurde.
    2. CSV-Import (Computername;MAC).
    3. Dynamische MAC-Ermittlung via DHCP, ARP oder SCCM.
    4. Subnetz-übergreifendes Wecken (Directed Broadcast).
    5. Optionale Nutzung der SCCM-Infrastruktur zum Wecken.
    6. "Smart-Retry": Prüft Erreichbarkeit und versucht es erneut, falls der PC nicht startet.

.PARAMETER Target
    Der Hostname ODER die MAC-Adresse des Ziels. 
    Das Skript unterscheidet automatisch anhand des Formats.

.PARAMETER CsvFile
    Pfad zu einer CSV-Datei mit Header "ComputerName;MacAddress" (Trennzeichen Semikolon für DE).

.PARAMETER BroadcastAddress
    Die Ziel-IP für den Broadcast. Standard: 255.255.255.255 (Lokales Subnetz).
    Für Routing über Subnetze hinweg: z.B. 192.168.10.255.

.PARAMETER Port
    Der UDP Port (Standard: 9).

.PARAMETER WaitForOnline
    Schalter: Prüft nach dem Senden, ob der Host online kommt (Ping).
    Falls nicht, wird der Weckvorgang wiederholt.

.PARAMETER RetryCount
    Anzahl der Wiederholungsversuche, falls -WaitForOnline aktiv ist (Standard: 3).

.PARAMETER RetryInterval
    Wartezeit in Sekunden zwischen den Versuchen (Standard: 30).

.PARAMETER UseSCCM
    Schalter: Versucht den Wake-Up Befehl über den SCCM Server zu senden (benötigt WMI Zugriff).

.PARAMETER SccmServer
    Der Name des SCCM Site Servers (nur benötigt bei -UseSCCM).

.PARAMETER SccmSiteCode
    Der SCCM Site Code (z.B. P01) (nur benötigt bei -UseSCCM).

.EXAMPLE
    .\Invoke-WOLManager.ps1 PC01 -WaitForOnline
    Weckt PC01, wartet, pingt. Wenn offline, erneuter Versuch (inkl. neuer MAC-Suche).

.NOTES
    File Name  : Invoke-WOLManager.ps1
    Author     : PS-Coding
    Version    : 1.3 (Added Connectivity Check & Smart Retry)
    Requires   : PowerShell 5.1
#>

[CmdletBinding(DefaultParameterSetName="Interactive")]
param(
    [Parameter(ParameterSetName="SmartInput", Mandatory=$true, Position=0, ValueFromPipeline=$true)]
    [Parameter(ParameterSetName="SCCM", Mandatory=$true, Position=0)]
    [Alias("ComputerName", "MacAddress", "HostName", "Identity")]
    [string[]]$Target,

    [Parameter(ParameterSetName="CsvImport", Mandatory=$true)]
    [string]$CsvFile,

    [Parameter(ParameterSetName="SmartInput")]
    [Parameter(ParameterSetName="CsvImport")]
    [string]$BroadcastAddress = "255.255.255.255",

    [int]$Port = 9,
    
    [Parameter(ParameterSetName="SmartInput")]
    [Parameter(ParameterSetName="CsvImport")]
    [Parameter(ParameterSetName="SCCM")]
    [switch]$WaitForOnline,

    [int]$RetryCount = 3,
    
    [int]$RetryInterval = 30,

    [Parameter(ParameterSetName="SCCM", Mandatory=$true)]
    [switch]$UseSCCM,

    [Parameter(ParameterSetName="SCCM")]
    [string]$SccmServer,

    [Parameter(ParameterSetName="SCCM")]
    [string]$SccmSiteCode
)

# ---------------------------------------------------------------------------
# Globale Einstellungen & Logging
# ---------------------------------------------------------------------------
$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "Info" # Info, Warning, Error, Verbose, Success
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $FormattedMsg = "[$Timestamp] [$Level] $Message"

    switch ($Level) {
        "Verbose" { Write-Verbose $Message }
        "Warning" { Write-Warning $FormattedMsg }
        "Error"   { Write-Host $FormattedMsg -ForegroundColor Red }
        "Success" { Write-Host $FormattedMsg -ForegroundColor Green }
        Default   { Write-Host $FormattedMsg -ForegroundColor Cyan }
    }
}

# ---------------------------------------------------------------------------
# Hilfsfunktionen
# ---------------------------------------------------------------------------

function Test-IsMacAddress {
    param([string]$InputString)
    return ($InputString -match "^([0-9a-fA-F]{2}[:-]){5}([0-9a-fA-F]{2})$|^[0-9a-fA-F]{12}$")
}

function Format-MacAddress {
    param([string]$Mac)
    $CleanMac = $Mac -replace "[^a-fA-F0-9]", ""
    if ($CleanMac.Length -ne 12) {
        throw "Ungültige MAC-Adresse: $Mac (Länge muss 12 Hex-Zeichen sein)"
    }
    return -join ($CleanMac.ToCharArray() | ForEach-Object { 
        $i = 0 
    } { 
        $_; if (++$i % 2 -eq 0 -and $i -lt 12) { ":" } 
    })
}

function Get-MacFromInfrastructure {
    param(
        [string]$TargetName,
        [switch]$TrySCCM,
        [string]$SiteServer,
        [string]$SiteCode
    )
    
    Write-Log "  -> Suche MAC-Adresse für '$TargetName'..." "Verbose"
    
    # 1. Versuch: ARP Cache (Refresh der Daten vor Abfrage)
    try {
        # Versuche DNS Auflösung für IP
        $IP = (Resolve-DnsName $TargetName -ErrorAction SilentlyContinue).IPAddress | Select-Object -First 1
        if ($IP) {
            $ArpEntry = Get-NetNeighbor -IPAddress $IP -ErrorAction SilentlyContinue
            if ($ArpEntry.LinkLayerAddress) {
                Write-Log "  -> MAC via ARP gefunden ($($ArpEntry.LinkLayerAddress))." "Verbose"
                return $ArpEntry.LinkLayerAddress
            }
        }
    } catch { Write-Log "ARP Lookup fehlgeschlagen." "Verbose" }

    # 2. Versuch: DHCP Server (Benötigt RSAT)
    try {
        if (Get-Command Get-DhcpServerv4Lease -ErrorAction SilentlyContinue) {
            $Lease = Get-DhcpServerv4Lease -ComputerName (Get-DhcpServerInDC).DnsName -ScopeId 0.0.0.0 | Where-Object { $_.HostName -like "$TargetName*" } | Select-Object -First 1
            if ($Lease) {
                Write-Log "  -> MAC via DHCP gefunden ($($Lease.ClientId))." "Verbose"
                return $Lease.ClientId
            }
        }
    } catch { Write-Log "DHCP Lookup nicht möglich/fehlgeschlagen." "Verbose" }

    # 3. Versuch: SCCM Datenbank
    if ($TrySCCM -and $SiteServer -and $SiteCode) {
        try {
            $WmiQuery = "SELECT MACAddresses FROM SMS_R_System WHERE Name = '$TargetName'"
            $SccmClient = Get-CimInstance -Namespace "root\SMS\site_$SiteCode" -ComputerName $SiteServer -Query $WmiQuery -ErrorAction Stop
            if ($SccmClient.MACAddresses) {
                Write-Log "  -> MAC via SCCM DB gefunden." "Verbose"
                return $SccmClient.MACAddresses[0]
            }
        } catch { Write-Log "SCCM DB Lookup fehlgeschlagen." "Verbose" }
    }

    return $null
}

function Send-MagicPacket {
    param(
        [string]$Mac,
        [string]$BroadcastIP,
        [int]$TargetPort
    )
    try {
        $MacClean = $Mac -replace "[:\-]", ""
        $ByteMac = New-Object byte[] 6
        $ByteMac = for ($i = 0; $i -lt $MacClean.Length; $i += 2) {
            [convert]::ToByte($MacClean.Substring($i, 2), 16)
        }
        $Header = [byte[]](,0xFF * 6)
        $Payload = [byte[]]($ByteMac * 16)
        $Packet = $Header + $Payload

        $UdpClient = New-Object System.Net.Sockets.UdpClient
        $UdpClient.EnableBroadcast = $true
        $Endpoint = [System.Net.IPEndPoint]::new([System.Net.IPAddress]::Parse($BroadcastIP), $TargetPort)
        [void]$UdpClient.Send($Packet, $Packet.Length, $Endpoint)
        $UdpClient.Close()
        
        Write-Log "  -> Magic Packet gesendet an $Mac ($BroadcastIP)." "Info"
        return $true
    }
    catch {
        Write-Log "  -> Fehler beim Senden: $($_.Exception.Message)" "Error"
        return $false
    }
}

function Invoke-SccmWaker {
    param(
        [string]$TargetName,
        [string]$Server,
        [string]$Code
    )
    try {
        Write-Log "  -> Initiiere SCCM Wake-Up für $TargetName..." "Info"
        $SystemResource = Get-CimInstance -Namespace "root\SMS\site_$Code" -ComputerName $Server -Query "SELECT * FROM SMS_R_System WHERE Name = '$TargetName'"
        if (-not $SystemResource) { throw "PC nicht im SCCM." }
        
        $Mac = $SystemResource.MACAddresses[0]
        $WmiClass = [wmiclass]"\\$Server\root\SMS\site_$($Code):SMS_Sleepyho"
        $Result = $WmiClass.TriggerWakeUp($Mac, $null, 2, $null)
        
        if ($Result.ReturnValue -eq 0) { return $true } else { return $false }
    }
    catch {
        Write-Log "  -> SCCM Fehler: $($_.Exception.Message)" "Error"
        return $false
    }
}

function Start-HostSmart {
    param(
        [string]$NameOrMac,
        [string]$ExplicitMac = $null,
        [bool]$UseSCCMMode,
        [int]$CurrentRetryCount
    )

    $LoopLimit = if ($WaitForOnline) { $CurrentRetryCount } else { 0 }
    $CurrentLoop = 0

    do {
        $CurrentLoop++
        if ($CurrentLoop -gt 1) { Write-Log "Wiederhole Vorgang (Versuch $CurrentLoop von $($LoopLimit + 1))..." "Warning" }

        # --- SCHRITT 1: Erkennung & Versand ---
        
        # A) Ist es eine reine MAC?
        if (Test-IsMacAddress $NameOrMac) {
            # Bei reiner MAC können wir nicht viel "neu ermitteln", wir senden einfach
            Send-MagicPacket -Mac (Format-MacAddress $NameOrMac) -BroadcastIP $BroadcastAddress -TargetPort $Port
        }
        # B) Ist es ein Hostname (oder SCCM Modus)?
        else {
            if ($UseSCCMMode) {
                # SCCM Versuch
                Invoke-SccmWaker -TargetName $NameOrMac -Server $SccmServer -Code $SccmSiteCode
            }
            else {
                # Standard Modus: MAC suchen
                # HIER ist die Logik für "Prüfung ob andere Wege zur Verfügung stehen"
                # Wir rufen Get-MacFromInfrastructure in JEDER Schleife neu auf.
                # Falls ARP eben leer war, aber jetzt da ist, wird es gefunden.
                
                $DynMac = if ($ExplicitMac) { $ExplicitMac } else { 
                    Get-MacFromInfrastructure -TargetName $NameOrMac -TrySCCM:$UseSCCM -SiteServer:$SccmServer -SiteCode:$SccmSiteCode 
                }

                if ($DynMac) {
                    Send-MagicPacket -Mac (Format-MacAddress $DynMac) -BroadcastIP $BroadcastAddress -TargetPort $Port
                } else {
                    Write-Log "  -> Keine MAC ermittelbar." "Warning"
                }
            }
        }

        # --- SCHRITT 2: Prüfung & Warten ---
        
        if ($WaitForOnline) {
            # Hinweis: Ping macht nur Sinn bei Hostnamen, nicht bei reiner MAC (ohne IP)
            if (Test-IsMacAddress $NameOrMac) {
                Write-Log "  -> Überspringe Ping-Check (Nur MAC Adresse bekannt)." "Verbose"
                break 
            }
            
            Write-Log "  -> Prüfe Erreichbarkeit (Ping an $NameOrMac)..." "Verbose"
            if (Test-Connection -ComputerName $NameOrMac -Count 1 -Quiet -ErrorAction SilentlyContinue) {
                Write-Log "$NameOrMac ist ONLINE!" "Success"
                return # Erfolg, raus aus der Schleife
            } else {
                if ($CurrentLoop -le $LoopLimit) {
                    Write-Log "  -> $NameOrMac antwortet noch nicht. Warte $RetryInterval Sekunden..."
                    Start-Sleep -Seconds $RetryInterval
                } else {
                    Write-Log "$NameOrMac ist nach $($LoopLimit + 1) Versuchen immer noch OFFLINE." "Error"
                }
            }
        }

    } while ($CurrentLoop -le $LoopLimit)
}

# ---------------------------------------------------------------------------
# Hauptprogramm
# ---------------------------------------------------------------------------

Write-Log "Starte WOL Manager..." "Verbose"

# INTERACTIVE MODE
if ($PSCmdlet.ParameterSetName -eq "Interactive") {
    Write-Log "Interaktiver Modus."
    $Mode = Read-Host "Modus wählen (1: Einzeln/Liste, 2: CSV, 3: SCCM)"
    switch ($Mode) {
        "1" { 
            $Target = @((Read-Host "Computername(n) oder MAC(s)")) 
            $BroadcastAddress = Read-Host "Broadcast IP (Leer lassen für 255.255.255.255)"
            if ([string]::IsNullOrWhiteSpace($BroadcastAddress)) { $BroadcastAddress = "255.255.255.255" }
            if ((Read-Host "Warten bis Online? (j/n)") -eq "j") { $WaitForOnline = $true }
        }
        "2" { $CsvFile = Read-Host "Pfad zur CSV Datei" }
        "3" {
            $UseSCCM = $true
            $Target = @((Read-Host "Computername eingeben"))
            $SccmServer = Read-Host "SCCM Server Name"
            $SccmSiteCode = Read-Host "SCCM Site Code"
        }
    }
}

# VERARBEITUNG

# Fall A: CSV Import
if ($CsvFile) {
    if (Test-Path $CsvFile) {
        Write-Log "Verarbeite CSV..." "Verbose"
        $Clients = Import-Csv -Path $CsvFile -Delimiter ";" -Encoding Default
        
        foreach ($Client in $Clients) {
            # Entscheide Input (Name oder MAC bevorzugt?)
            # Wir übergeben Name wenn da, damit Ping Check klappt. MAC explizit wenn da.
            if ($Client.ComputerName) {
                Start-HostSmart -NameOrMac $Client.ComputerName -ExplicitMac $Client.MacAddress -UseSCCMMode $false -CurrentRetryCount $RetryCount
            } elseif ($Client.MacAddress) {
                Start-HostSmart -NameOrMac $Client.MacAddress -UseSCCMMode $false -CurrentRetryCount $RetryCount
            }
        }
    } else { Write-Log "CSV Datei nicht gefunden!" "Error" }
}

# Fall B & C: Direktes Ziel (SCCM oder Smart)
elseif ($Target) {
    foreach ($Item in $Target) {
        Start-HostSmart -NameOrMac $Item -UseSCCMMode $UseSCCM -CurrentRetryCount $RetryCount
    }
}

Write-Log "Vorgang abgeschlossen." "Info"
