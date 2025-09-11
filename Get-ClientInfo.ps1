<#
.SYNOPSIS
    Retrieves hostname, MAC address, and IP address for computers matching a specified name pattern.
    The script attempts to get this information for both online and offline clients.

.DESCRIPTION
    This script queries Active Directory (if available) for computer objects that match a given name or wildcard pattern.
    It then iterates through each found computer.
    - Hostname is retrieved from Active Directory.
    - MAC Address is attempted to be retrieved using a remote WMI/CIM call. This will only work for online computers.
    - IP Address is attempted to be retrieved via DNS lookup after a successful network ping.

.PARAMETER ComputerName
    A name or a wildcard pattern for the computers to search for.
    If no value is provided, the script will run in interactive mode and prompt the user.

.PARAMETER OutputPath
    The full path to the output CSV file. Defaults to "C:\Temp\ClientInfo.csv".

.PARAMETER VerboseOutput
    Enables verbose output to display detailed execution information. Enabled by default.

.EXAMPLE
    PS C:\> .\Get-ClientInfo.ps1

    Runs the script in interactive mode, prompting for the computer name.

.EXAMPLE
    PS C:\> .\Get-ClientInfo.ps1 -ComputerName "WS-*"

    Searches for all computers with a hostname starting with "WS-" and saves the output to the default path.

.EXAMPLE
    PS C:\> .\Get-ClientInfo.ps1 -ComputerName "Laptop-*" -OutputPath "C:\Users\User\Documents\Laptops.csv"

    Searches for all computers with a hostname starting with "Laptop-" and saves the output to a custom path.

.EXAMPLE
    PS C:\> .\Get-ClientInfo.ps1 -ComputerName "SRV-DB01" -VerboseOutput:$false

    Searches for a specific server and suppresses detailed verbose output.

.NOTES
    - Author: PS-Coding
    - Version: 1.0
    - Creation Date: 2025-09-11
    - Last Modified: 2025-09-11
    - Localization: German-specific settings are applied (CSV delimiter ';').
    - Source References:
      - AI-discovered:
        - How to query Active Directory with PowerShell
        - How to get MAC address remotely with PowerShell
        - Exporting to CSV with custom delimiters
#>

# --- Configuration Section ---
# This section defines the script parameters and default values.
[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = "Interactive")]
param (
    [Parameter(Mandatory = $false, HelpMessage = "The computer name or wildcard to search for.")]
    [string]$ComputerName,

    [Parameter(Mandatory = $false, HelpMessage = "The full path for the output CSV file.")]
    [string]$OutputPath = "C:\Temp\ClientInfo.csv",

    [Parameter(Mandatory = $false, HelpMessage = "Enable verbose output.")]
    [switch]$VerboseOutput = $true
)

# --- Dynamic Environment Discovery ---
# This section dynamically determines necessary environment data.
Write-Verbose -Verbose:$VerboseOutput "Initialisiere Skript und prüfe Umgebung..."

# Get the current domain name to query Active Directory.
try {
    $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name
    Write-Verbose -Verbose:$VerboseOutput "Erfolgreich die Domäne erkannt: '$domain'."
    $activeDirectoryModule = "ActiveDirectory"
} catch {
    Write-Warning "Konnte die Active Directory Domäne nicht erkennen. Das Skript funktioniert möglicherweise nur für Computer im lokalen Netzwerk oder erfordert manuelle Eingabe von Computernamen."
    $activeDirectoryModule = $null
}

# --- Main Script Body ---
# This is the core logic of the script.

# Function to get the list of computers
function Get-Computers {
    Write-Verbose -Verbose:$VerboseOutput "Ermittle Computernamen basierend auf dem Suchmuster '$ComputerName'..."
    $computers = @()

    if ($activeDirectoryModule) {
        # Import the ActiveDirectory module if available.
        if (-not (Get-Module -Name $activeDirectoryModule -ListAvailable)) {
            Write-Warning "Das Modul '$activeDirectoryModule' wurde nicht gefunden. Versuche es zu installieren oder zu importieren."
        }
        Import-Module $activeDirectoryModule -ErrorAction SilentlyContinue

        # Get computer objects from Active Directory.
        $computers = Get-ADComputer -Filter "Name -like '$ComputerName'" -Properties DNSHostName | Select-Object -ExpandProperty DNSHostName
        Write-Verbose -Verbose:$VerboseOutput "Gefundene Computer im Active Directory: $($computers.Count)"
    } else {
        # Fallback for non-AD environments.
        Write-Verbose -Verbose:$VerboseOutput "Keine Active Directory Verbindung. Suche Computer mit dem ping-Befehl."
        Write-Warning "Die Suche ist auf erreichbare Computer beschränkt."
        
        # This part is a simplified fallback; for a robust solution in non-AD environments,
        # a different approach (like a list of known hosts) would be needed.
        # For now, we assume the user provides full names in a non-AD context.
        $computers = $ComputerName
    }

    if ($null -eq $computers -or $computers.Count -eq 0) {
        Write-Warning "Keine Computer mit dem Namen '$ComputerName' gefunden. Skript wird beendet."
        return $null
    }
    return $computers
}

# Function to process each computer and get its information.
function Get-ComputerInfo {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Hostname
    )

    Write-Verbose -Verbose:$VerboseOutput "Verarbeite Hostname: '$Hostname'..."
    $infoObject = New-Object -TypeName PSObject -Property @{
        Hostname    = $Hostname
        IPAddress   = "Offline"
        MACAddress  = "Offline"
    }

    # Check if the computer is online.
    try {
        $pingResult = Test-Connection -ComputerName $Hostname -Count 1 -ErrorAction Stop
        if ($pingResult) {
            Write-Verbose -Verbose:$VerboseOutput "Client ist online. Versuche, die IP- und MAC-Adresse zu ermitteln."

            # Get IP Address from DNS.
            try {
                $dnsLookup = Resolve-DnsName -Name $Hostname -ErrorAction Stop
                $infoObject.IPAddress = $dnsLookup.IPAddress.Where({ $_.Length -le 15 }, 'First')
                Write-Verbose -Verbose:$VerboseOutput "IP-Adresse gefunden: $($infoObject.IPAddress)"
            } catch {
                Write-Warning "Konnte die IP-Adresse für '$Hostname' nicht ermitteln."
            }

            # Get MAC Address using CIM.
            try {
                $macAddress = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -ComputerName $Hostname -ErrorAction Stop |
                              Where-Object { $_.IPEnabled } | Select-Object -ExpandProperty MACAddress
                if ($macAddress) {
                    $infoObject.MACAddress = $macAddress
                    Write-Verbose -Verbose:$VerboseOutput "MAC-Adresse gefunden: $($infoObject.MACAddress)"
                } else {
                    Write-Warning "Konnte die MAC-Adresse für '$Hostname' nicht ermitteln."
                }
            } catch {
                Write-Warning "Konnte die MAC-Adresse für '$Hostname' nicht ermitteln. Möglicherweise fehlen Zugriffsrechte."
            }
        }
    } catch {
        Write-Verbose -Verbose:$VerboseOutput "Client '$Hostname' ist offline oder nicht erreichbar."
    }

    return $infoObject
}

# --- Execution Flow ---
if ($PSCmdlet.ParameterSetName -eq "Interactive") {
    # Interactive mode: Prompt for the computer name.
    Write-Host "Bitte geben Sie einen Computernamen oder ein Wildcard-Muster ein (z.B. 'WS-*', 'SRV-DB01'):"
    $ComputerName = Read-Host
}

# Pre-processing and Confirmation
Write-Host "Das Skript wird die Suche nach '$ComputerName' starten und die Ergebnisse in '$OutputPath' speichern."
$choice = Read-Host "Möchten Sie fortfahren? (J/N)"
if ($choice -ne "J" -and $choice -ne "j") {
    Write-Warning "Skript-Ausführung abgebrochen."
    exit
}

# Create output folder if it doesn't exist.
$outputDir = Split-Path -Path $OutputPath -Parent
if (-not (Test-Path -Path $outputDir)) {
    Write-Verbose -Verbose:$VerboseOutput "Erstelle das Ausgabeverzeichnis: '$outputDir'."
    New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
}

$results = @()
$computers = Get-Computers

if ($computers) {
    foreach ($computer in $computers) {
        $results += Get-ComputerInfo -Hostname $computer
    }
}

# Export results to CSV with German delimiter
if ($results.Count -gt 0) {
    Write-Verbose -Verbose:$VerboseOutput "Speichere die Ergebnisse in '$OutputPath'..."
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Delimiter ';' -Encoding UTF8
    Write-Host "Vorgang abgeschlossen! Die Ergebnisse wurden in '$OutputPath' gespeichert."
} else {
    Write-Warning "Keine Ergebnisse zum Speichern gefunden. Die CSV-Datei wurde nicht erstellt."
}
