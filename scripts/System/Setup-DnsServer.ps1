<#
.SYNOPSIS
    Installs, configures, and repairs DNS servers and zones on Windows Server.

.DESCRIPTION
    This script consolidates dns-schul.ps1 and DNS_repair.ps1. It can install the DNS Server feature,
    create and configure Forward and Reverse Lookup Zones, configure forwarders, manage resource records,
    and perform zone reload or repair actions (e.g., reset zone type to Directory Integrated Primary).

.PARAMETER Install
    Switch to install the Windows DNS Server feature and administration tools if not present.

.PARAMETER SetupZones
    Switch to create or update Forward and Reverse Lookup zones.

.PARAMETER ConfigureForwarders
    Switch to set DNS forwarders.

.PARAMETER RepairZone
    Switch to reload a zone and reset its type (e.g., to Directory Integrated Primary).

.PARAMETER ZoneName
    The name of the Forward Lookup Zone (e.g., 'example.local').

.PARAMETER ReverseNetworkId
    The network ID for the Reverse Lookup Zone (e.g., '192.168.250.0/24').

.PARAMETER Forwarders
    An array of IP addresses for external DNS forwarders (e.g. @('1.1.1.1', '8.8.8.8')).

.PARAMETER HostName
    The name of a host A record to create/update.

.PARAMETER HostIP
    The IPv4 address of the host record to create/update.

.NOTES
    Version: 1.0
    Autor: Systems Administration
    Compatibility: PowerShell 5.1+
#>

[CmdletBinding()]
param (
    [Parameter()]
    [switch]$Install,

    [Parameter()]
    [switch]$SetupZones,

    [Parameter()]
    [switch]$ConfigureForwarders,

    [Parameter()]
    [switch]$RepairZone,

    [Parameter()]
    [string]$ZoneName = "lafp.schul.polizei.local",

    [Parameter()]
    [string]$ReverseNetworkId = "192.168.250.0/24",

    [Parameter()]
    [string[]]$Forwarders = @("1.1.1.1", "8.8.8.8"),

    [Parameter()]
    [string]$HostName = "SDOM",

    [Parameter()]
    [string]$HostIP = "192.168.250.203"
)

# Load environment configuration if available in parent directories
$envLoader = Join-Path $PSScriptRoot "..\..\Import-Environment.ps1"
if (Test-Path $envLoader) {
    . $envLoader
}

# 1. Install DNS Role
if ($Install) {
    Write-Host "Checking/Installing DNS server role..." -ForegroundColor Cyan
    try {
        $null = Install-WindowsFeature -Name DNS -IncludeManagementTools -ErrorAction Stop
        Write-Host "DNS server role check/install completed." -ForegroundColor Green
    } catch {
        Write-Error "Failed to install DNS server role: $_"
        return
    }
}

# Start DNS Service
if ($Install -or $SetupZones -or $RepairZone -or $ConfigureForwarders) {
    Write-Verbose "Ensuring DNS service is running..."
    try {
        $dnsService = Get-Service -Name DNS -ErrorAction Stop
        if ($dnsService.Status -ne "Running") {
            Start-Service -Name DNS -ErrorAction Stop
            Write-Host "DNS service started." -ForegroundColor Green
        }
    } catch {
        Write-Error "DNS Service is not available or failed to start: $_"
        return
    }
}

# 2. Setup Zones
if ($SetupZones) {
    Write-Host "Configuring Forward Lookup Zone '$ZoneName'..." -ForegroundColor Cyan
    try {
        if (-not (Get-DnsServerZone -Name $ZoneName -ErrorAction SilentlyContinue)) {
            $null = Add-DnsServerPrimaryZone -Name $ZoneName -ZoneFile "$ZoneName.dns" -DynamicUpdate Secure -ErrorAction Stop
            Write-Host "Created new Forward Zone: $ZoneName" -ForegroundColor Green
        } else {
            $null = Set-DnsServerPrimaryZone -Name $ZoneName -DynamicUpdate Secure -ErrorAction Stop
            Write-Host "Configured dynamic updates for Forward Zone: $ZoneName" -ForegroundColor Green
        }
    } catch {
        Write-Error "Error configuring Forward Lookup Zone: $_"
    }

    # Determine Reverse Zone Name from NetworkId
    # Example NetworkId: 192.168.250.0/24
    if ($ReverseNetworkId -match '^(\d+)\.(\d+)\.(\d+)\.0/24$') {
        $reverseZoneName = "$($Matches[3]).$($Matches[2]).$($Matches[1]).in-addr.arpa"
        Write-Host "Configuring Reverse Lookup Zone '$reverseZoneName'..." -ForegroundColor Cyan
        try {
            if (-not (Get-DnsServerZone -Name $reverseZoneName -ErrorAction SilentlyContinue)) {
                $null = Add-DnsServerPrimaryZone -NetworkId $ReverseNetworkId -ZoneFile "$reverseZoneName.dns" -DynamicUpdate Secure -ErrorAction Stop
                Write-Host "Created new Reverse Zone: $reverseZoneName" -ForegroundColor Green
            } else {
                $null = Set-DnsServerPrimaryZone -Name $reverseZoneName -DynamicUpdate Secure -ErrorAction Stop
                Write-Host "Configured dynamic updates for Reverse Zone: $reverseZoneName" -ForegroundColor Green
            }
        } catch {
            Write-Error "Error configuring Reverse Lookup Zone: $_"
        }
    } else {
        Write-Warning "ReverseNetworkId format is not in standard C-Class subnet format (X.Y.Z.0/24). Skipping Reverse Zone creation."
    }
}

# 3. Configure Forwarders
if ($ConfigureForwarders) {
    Write-Host "Configuring DNS forwarders: $($Forwarders -join ', ')..." -ForegroundColor Cyan
    try {
        $null = Set-DnsServerForwarder -IPAddress $Forwarders -ErrorAction Stop
        Write-Host "DNS forwarders set successfully." -ForegroundColor Green
    } catch {
        Write-Error "Failed to set DNS forwarders: $_"
    }
}

# 4. Repair Zone
if ($RepairZone) {
    Write-Host "Performing repair on zone '$ZoneName'..." -ForegroundColor Cyan
    try {
        # Reload zone
        $null = dnscmd /ZoneReload $ZoneName
        Write-Host "Zone reloaded successfully." -ForegroundColor Green
        
        # Reset zone type to Directory Integrated Primary
        $null = dnscmd /ZoneResetType $ZoneName /DsPrimary
        Write-Host "Zone type reset to Active Directory-integrated primary." -ForegroundColor Green
    } catch {
        Write-Error "Failed to repair zone '$ZoneName': $_"
    }
}

# 5. Manage Resource Record
if ($HostName -and $HostIP) {
    Write-Host "Ensuring Host A record '$HostName' points to '$HostIP' in zone '$ZoneName'..." -ForegroundColor Cyan
    try {
        $record = Get-DnsServerResourceRecord -ZoneName $ZoneName -Name $HostName -ErrorAction SilentlyContinue
        if (-not $record) {
            $null = Add-DnsServerResourceRecordA -ZoneName $ZoneName -Name $HostName -IPv4Address $HostIP -ErrorAction Stop
            Write-Host "Created new Host record A: $HostName -> $HostIP" -ForegroundColor Green
        } else {
            Write-Verbose "Record exists. Updating to new IP..."
            # Check if current IP differs
            if ($record.RecordData.IPv4Address.IPAddressToString -ne $HostIP) {
                # Remove old record and add new record (Powershell 5.1 safe update method)
                $null = Remove-DnsServerResourceRecord -ZoneName $ZoneName -InputObject $record -Force -ErrorAction Stop
                $null = Add-DnsServerResourceRecordA -ZoneName $ZoneName -Name $HostName -IPv4Address $HostIP -ErrorAction Stop
                Write-Host "Updated Host record A: $HostName -> $HostIP" -ForegroundColor Green
            } else {
                Write-Host "Host record A is already up-to-date: $HostName -> $HostIP" -ForegroundColor Green
            }
        }
    } catch {
        Write-Error "Error managing host A record: $_"
    }
}

Write-Host "DNS Setup/Repair operation completed." -ForegroundColor Cyan
