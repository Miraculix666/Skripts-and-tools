#
# FILE: Start-RemoteSessions.ps1
# DESCRIPTION: Opens multiple remote sessions, tiles them on the second monitor, and provides a comprehensive connectivity test.
#
# VERSION: 1.4.0
# DATE: 2025-08-28
# AUTHOR: PS-Coding
#
<#
.SYNOPSIS
    Starts multiple RDP sessions, tiles them on the second monitor, and handles credentials flexibly.
    Includes an optional, parallel connectivity test (Ping, RDP Port, and DNS).

.DESCRIPTION
    This script reads a list of computer names from an internal array, a text file, or a CSV file.
    It supports multiple predefined internal lists via the -Preset parameter.
    It then opens a new Remote Desktop Connection (mstsc.exe) for each computer, placing the windows
    in a tiled layout on the second display. Credentials can be provided via parameters,
    interactively, or by modifying the script variables.
    A new feature allows for a comprehensive, parallel connectivity test to check host, RDP port, and DNS resolution.

.PARAMETER ComputerName
    A comma-separated list of computer names to connect to. This parameter is ignored if -InputFile is used.

.PARAMETER InputFile
    The full path to a text (.txt) or CSV (.csv) file containing the list of computer names.
    For a text file, each computer name should be on a new line.
    For a CSV file, the column with computer names must be named "ComputerName".

.PARAMETER Preset
    The name(s) of one or more predefined internal computer lists to use. Available presets are defined
    in the script's internal variable. Examples: 'Default', 'R303', 'Server'. Multiple presets can be specified.

.PARAMETER UserName
    The username to use for the RDP connections.
    If not provided, the script will use the currently logged-in user's name.

.PARAMETER Password
    The password for the specified user. This is a SecureString.
    If not provided, the user will be prompted to enter it interactively.

.PARAMETER RDPResolution
    The resolution for the RDP sessions. Default is '1024,768'.
    Example: '1920,1080'.

.PARAMETER TestConnection
    Enables a connectivity test (Ping and RDP Port) before attempting to open RDP sessions.

.PARAMETER TestConnectionOnly
    Performs only the connectivity test and exits the script without opening any RDP sessions.
    This parameter takes precedence over -TestConnection.

.PARAMETER Verbose
    Enables detailed output during script execution. By default, verbose output is enabled.
    Can be switched off by specifying `-Verbose:$false`.

.EXAMPLE
    # Starts RDP sessions for the 'R303' and 'Server' preset lists with a prior connectivity test.
    .\Start-RemoteSessions.ps1 -Preset "R303", "Server" -TestConnection

.EXAMPLE
    # Performs a standalone connectivity test for the 'Server' and 'Testumgebung' preset lists.
    .\Start-RemoteSessions.ps1 -Preset "Server", "Testumgebung" -TestConnectionOnly

.EXAMPLE
    # Starts RDP sessions for computers in a text file without a connectivity test.
    .\Start-RemoteSessions.ps1 -InputFile C:\temp\computers.txt

#>
[CmdletBinding()]
param (
    [Parameter(Position = 0)]
    [string[]]$ComputerName = @(),

    [Parameter(Position = 1)]
    [string]$InputFile,

    [Parameter(Position = 2)]
    [string[]]$Preset,

    [Parameter(Position = 3)]
    [string]$UserName,

    [Parameter(Position = 4)]
    [System.Security.SecureString]$Password,

    [Parameter(Position = 5)]
    [string]$RDPResolution = "1024,768",
    
    [switch]$TestConnection,

    [switch]$TestConnectionOnly
)

# Set German locale for consistent formatting
Set-Culture -Culture de-DE

# --- SCRIPT-START ---

#region Function Definition
<#
.SYNOPSIS
    Tests if a computer is reachable, has the RDP port open, and its DNS name can be resolved.

.DESCRIPTION
    This function performs a parallel test for network reachability (ping), RDP port (3389)
    connectivity, and DNS resolution for a given list of computers.

.PARAMETER Computers
    An array of computer names or IP addresses to test.

.EXAMPLE
    Test-RDPConnection -Computers @("Server01", "PC-Test")
#>
function Test-RDPConnection {
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$Computers
    )

    Write-Verbose "Starte parallelen Konnektivitätstest für $($Computers.Count) Computer..."
    
    $results = @()
    $jobs = @()

    foreach ($computer in $Computers) {
        $job = Start-Job -ScriptBlock {
            param($computer)
            
            $pingStatus = "Unbekannt"
            $rdpPortStatus = "Unbekannt"
            $dnsStatus = "Unbekannt"
            
            # DNS Resolution Test
            try {
                $dnsResult = [System.Net.Dns]::GetHostAddresses($computer)
                $dnsStatus = "Aufgelöst"
            }
            catch {
                $dnsStatus = "Nicht aufgelöst"
            }
            
            # Ping test
            try {
                $pingResult = Test-Connection -ComputerName $computer -Count 1 -ErrorAction Stop
                $pingStatus = if ($pingResult.StatusCode -eq 0) { "Erreichbar" } else { "Nicht erreichbar" }
            }
            catch {
                $pingStatus = "Nicht erreichbar"
            }
            
            # RDP Port Test
            try {
                $tcpClient = New-Object System.Net.Sockets.TcpClient
                $connect = $tcpClient.BeginConnect($computer, 3389, $null, $null)
                $wait = $connect.AsyncWaitHandle.WaitOne(2000, $false)
                $rdpPortStatus = if ($wait -and $tcpClient.Connected) { "Geöffnet" } else { "Geschlossen" }
            }
            catch {
                $rdpPortStatus = "Geschlossen"
            }
            finally {
                if ($tcpClient) {
                    $tcpClient.Close()
                }
            }
            
            return [PSCustomObject]@{
                ComputerName  = $computer
                DNSStatus     = $dnsStatus
                PingStatus    = $pingStatus
                RDPPortStatus = $rdpPortStatus
            }
        } -ArgumentList $computer
        $jobs += $job
    }

    $jobs | Wait-Job | Receive-Job | ForEach-Object { $results += $_ }
    $jobs | Remove-Job -Force

    Write-Host "---"
    Write-Host "Konnektivitätsbericht:"
    $results | Format-Table -AutoSize
    Write-Host "---"
    
    return $results
}
#endregion

#region Variable and Parameter Definitions
Write-Verbose "Starte Skriptausführung am $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')."
Write-Verbose "---"
Write-Verbose "Parameter:"
Write-Verbose "  ComputerName:       '$($ComputerName -join ', ')'"
Write-Verbose "  InputFile:          '$InputFile'"
Write-Verbose "  Preset:             '$($Preset -join ', ')'"
Write-Verbose "  UserName:           '$UserName'"
Write-Verbose "  RDPResolution:      '$RDPResolution'"
Write-Verbose "  TestConnection:     '$TestConnection'"
Write-Verbose "  TestConnectionOnly: '$TestConnectionOnly'"
Write-Verbose "---"

# Default internal computer lists (as a Hashtable)
$InternalLists = @{
    "Default" = @(
        "Server01",
        "Server02",
        "Server03"
    );
    "R303" = @(
        "pc-r303-1",
        "pc-r303-2",
        "pc-r303-3",
        "pc-r303-4"
    );
    "Server" = @(
        "web-server01",
        "sql-server02"
    );
    "Testumgebung" = @(
        "dev-vm-01",
        "dev-vm-02",
        "test-vm-01"
    );
    "R202" = @(
        "C26PBZE70004200",
        "C26PBZE70004201",
        "C26PBZE70004202",
        "C26PBZE70004203",
        "C26PBZE70004204",
        "C26PBZE70004205",
        "C26PBZE70004206",
        "C26PBZE70004207",
        "C26PBZE70004208",
        "C26PBZE70004209",
        "C26PBZE70004210",
        "C26PBZE70004211",
        "C26PBZE70004212",
        "C26PBZE70004299"
    );
    "R204" = @(
        "C26PBZE70004400",
        "C26PBZE70004401",
        "C26PBZE70004402",
        "C26PBZE70004403",
        "C26PBZE70004404",
        "C26PBZE70004405",
        "C26PBZE70004406",
        "C26PBZE70004407",
        "C26PBZE70004408",
        "C26PBZE70004409",
        "C26PBZE70004410",
        "C26PBZE70004411",
        "C26PBZE70004412"
    );
    "R206" = @(
        "C26PBZE70004600",
        "C26PBZE70004601",
        "C26PBZE70004602",
        "C26PBZE70004603",
        "C26PBZE70004604",
        "C26PBZE70004605",
        "C26PBZE70004606",
        "C26PBZE70004607",
        "C26PBZE70004608",
        "C26PBZE70004609",
        "C26PBZE70004610",
        "C26PBZE70004611",
        "C26PBZE70004612"
    );    
    "R212" = @(
      "C26PBZE70004100",
      "C26PBZE70004101",
      "C26PBZE70004102",
      "C26PBZE70004103",
      "C26PBZE70004104",
      "C26PBZE70004105",
      "C26PBZE70004106",
      "C26PBZE70004107",
      "C26PBZE70004108",
      "C26PBZE70004109",
      "C26PBZE70004110",
      "C26PBZE70004111",
      "C26PBZE70004112",
      "C26PBZE70004113",
      "C26PBZE70004114",
      "C26PBZE70004115",
      "C26PBZE70004116"
    )
}
#endregion

#region Get Computer List
$ComputersToConnect = @()

if ($InputFile) {
    Write-Verbose "InputFile-Parameter angegeben. Lese Computernamen aus '$InputFile'."
    
    if (-not (Test-Path -Path $InputFile)) {
        Write-Error "Fehler: Die angegebene Eingabedatei '$InputFile' existiert nicht."
        return
    }

    $fileExtension = [System.IO.Path]::GetExtension($InputFile).ToLower()
    
    if ($fileExtension -eq ".csv") {
        try {
            $ComputersToConnect = Import-Csv -Path $InputFile -Delimiter ';' | Select-Object -ExpandProperty ComputerName
        }
        catch {
            Write-Error "Fehler beim Lesen der CSV-Datei."
            return
        }
    }
    elseif ($fileExtension -eq ".txt") {
        $ComputersToConnect = Get-Content -Path $InputFile | Where-Object { $_.Trim() -ne "" }
    }
    else {
        Write-Error "Nicht unterstütztes Dateiformat '$fileExtension'."
        return
    }
}
elseif ($ComputerName) {
    Write-Verbose "ComputerName-Parameter angegeben. Verwende die bereitgestellte Liste."
    $ComputersToConnect = $ComputerName
}
elseif ($Preset) {
    Write-Verbose "Preset-Parameter angegeben. Kombiniere die Listen: '$($Preset -join ', ')'."
    $allComputersFromPresets = @()
    foreach ($p in $Preset) {
        if ($InternalLists.ContainsKey($p)) {
            $allComputersFromPresets += $InternalLists[$p]
        }
        else {
            Write-Warning "Das angegebene Preset '$p' existiert nicht und wird übersprungen. Verfügbare Presets: $($InternalLists.Keys -join ', ')."
        }
    }
    $ComputersToConnect = $allComputersFromPresets | Sort-Object -Unique
}
else {
    Write-Verbose "Keine Parameter angegeben. Verwende die interne 'Default'-Liste."
    $ComputersToConnect = $InternalLists["Default"]
}

if (-not $ComputersToConnect) {
    Write-Warning "Keine Computer zum Verbinden gefunden. Skript wird beendet."
    return
}

Write-Verbose "Es wurden $($ComputersToConnect.Count) Computer zum Verbinden gefunden."

#endregion

#region Connectivity Test
if ($TestConnection -or $TestConnectionOnly) {
    $testResults = Test-RDPConnection -Computers $ComputersToConnect
    $failedTests = $testResults | Where-Object { $_.DNSStatus -eq "Nicht aufgelöst" -or $_.PingStatus -eq "Nicht erreichbar" -or $_.RDPPortStatus -eq "Geschlossen" }
    
    if ($TestConnectionOnly) {
        Write-Host "---"
        Write-Verbose "Skriptausführung beendet, da der -TestConnectionOnly-Parameter angegeben wurde."
        return
    }

    if ($failedTests.Count -gt 0) {
        Write-Warning "Es wurden $($failedTests.Count) Computer gefunden, die nicht erreichbar sind oder deren DNS-Namen nicht aufgelöst werden konnten."
        Write-Host "Möchtest du trotzdem versuchen, die RDP-Sitzungen zu öffnen? (J/N)" -NoNewline
        $continue = Read-Host
        if ($continue -notmatch "^[jJ]$") {
            Write-Host "RDP-Sitzungen werden nicht geöffnet. Skript wird beendet."
            return
        }
    }
}
#endregion

#region Handle Credentials
if (-not $UserName) {
    $UserName = $env:USERNAME
    Write-Verbose "Kein Benutzername angegeben, verwende den aktuellen Benutzer: '$UserName'."
}

if (-not $Password) {
    Write-Host "Bitte das Passwort für '$UserName' eingeben:" -NoNewline
    $Password = Read-Host -AsSecureString
    if (-not $Password) {
        Write-Error "Das Passwort darf nicht leer sein. Skript wird beendet."
        return
    }
}

#endregion

#region Get Screen Dimensions for Tiling
Write-Verbose "Bildschirmkonfiguration wird erkannt..."
try {
    # Dynamische Erkennung des zweiten Bildschirms
    $secondaryMonitor = Get-CimInstance -ClassName Win32_DesktopMonitor | Where-Object { $_.Availability -eq 3 -and $_.IsPrimary -ne $true } | Select-Object -First 1

    if ($null -eq $secondaryMonitor) {
        Write-Warning "Es wurde kein zweiter aktiver Bildschirm erkannt. Die RDP-Fenster werden normal und nicht gekachelt geöffnet."
        $useTiling = $false
    }
    else {
        $xOffset = [int]($secondaryMonitor.Left)
        $yOffset = [int]($secondaryMonitor.Top)
        $secondScreen_width = [int]($secondaryMonitor.ScreenWidth)
        $secondScreen_height = [int]($secondaryMonitor.ScreenHeight)
    
        Write-Verbose "Zweiter Bildschirm erkannt. Positionierung bei X-Offset: ${xOffset}px."
        $useTiling = $true
    }
}
catch {
    Write-Error "Fehler beim Abrufen der Monitorinformationen. Das Kacheln wird nicht verwendet."
    Write-Error "Fehlermeldung: $($_.Exception.Message)"
    $useTiling = $false
}

#endregion

#region Create and Start RDP Sessions
Write-Verbose "RDP-Sitzungen werden gestartet..."

# Dynamic tiling logic
if ($useTiling) {
    $rdpSize = $RDPResolution.Split(',') | ForEach-Object { [int]$_ }
    $rdpWidth = $rdpSize[0]
    $rdpHeight = $rdpSize[1]
    
    $numComputers = $ComputersToConnect.Count
    $numRows = [System.Math]::Ceiling([System.Math]::Sqrt($numComputers))
    $numCols = [System.Math]::Ceiling($numComputers / $numRows)
}

$counter = 0
foreach ($computer in $ComputersToConnect) {
    $counter++
    Write-Verbose "Verarbeite Computer $counter/$($ComputersToConnect.Count): '$computer'..."
    
    # Create a temporary RDP file
    $tempRDPFile = Join-Path -Path $env:TEMP -ChildPath "$computer-RDP.rdp"
    
    $rdpContent = @"
redirectclipboard:i:1
redirectprinters:i:1
redirectcomports:i:1
redirectsmartcards:i:1
devicestoredirect:s:*
drivestoredirect:s:*
full address:s:$computer
server port:i:3389
username:s:$UserName
desktopwidth:i:$rdpWidth
desktopheight:i:$rdpHeight
session bpp:i:24
audiomode:i:2
connection type:i:7
promptcredentialonce:i:0
autoreconnect:i:1
"@
    
    if ($useTiling) {
        $row = [int](($counter - 1) / $numCols)
        $col = [int](($counter - 1) % $numCols)
        $xPos = $xOffset + ($col * $rdpWidth)
        $yPos = $yOffset + ($row * $rdpHeight)
        
        $tilingContent = "winposstr:s:0,0,${yPos},${xPos},$($rdpHeight),$($rdpWidth)"
        $rdpContent += "`n$tilingContent"
    }
    
    $rdpContent += "`nfullscreen:i:0"

    Write-Verbose "  RDP-Datei wird erstellt unter: $tempRDPFile"
    Set-Content -Path $tempRDPFile -Value $rdpContent
    
    Write-Verbose "  Versuche, RDP-Sitzung für '$computer' zu öffnen..."
    Start-Process -FilePath "mstsc.exe" -ArgumentList "$tempRDPFile" -PassThru -NoNewWindow
    
    Start-Sleep -Milliseconds 500
}

Write-Verbose "Alle RDP-Sitzungen wurden gestartet."
Write-Host "---"
Write-Host "Zusammenfassung:"
Write-Host "  Computer:  $($ComputersToConnect.Count)"
Write-Host "  Auflösung: $RDPResolution"
Write-Host "  Benutzer:   $UserName"
Write-Host "---"

# End of script
Write-Verbose "Skriptausführung beendet."
```
eof

### **Anleitung zur Verwendung**

Um die neuen Funktionen zu nutzen, führe das Skript wie folgt aus:

* **Mehrere Presets kombinieren**: Gib einfach alle gewünschten Preset-Namen durch Kommas getrennt an. Das Skript erkennt die Namen und kombiniert die Listen automatisch.
    ```powershell
    .\Start-RemoteSessions.ps1 -Preset "R303", "R204", "R212"
    ```
* **Nur die Konnektivität prüfen**: Um die PCs schnell zu überprüfen, ohne eine RDP-Sitzung zu starten, nutze den Parameter **`-TestConnectionOnly`**.
    ```powershell
    .\Start-RemoteSessions.ps1 -Preset "R202" -TestConnectionOnly
    
