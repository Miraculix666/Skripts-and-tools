# Variablen definieren
$zoneName = "example.local"
$reverseZoneName = "250.168.192.in-addr.arpa"
$dnsServer = "localhost"
$forwarders = @("1.1.1.1", "8.8.8.8") # Cloudflare und Google DNS

# DNS-Serverrolle installieren, falls nicht vorhanden
Install-WindowsFeature -Name DNS -IncludeManagementTools

# DNS-Serverdienst starten
Start-Service -Name DNS

# Forward Lookup Zone erstellen oder aktualisieren
if (-not (Get-DnsServerZone -Name $zoneName -ErrorAction SilentlyContinue)) {
    Add-DnsServerPrimaryZone -Name $zoneName -ZoneFile "$zoneName.dns" -DynamicUpdate Secure
} else {
    Set-DnsServerPrimaryZone -Name $zoneName -DynamicUpdate Secure
}

# Reverse Lookup Zone erstellen oder aktualisieren
if (-not (Get-DnsServerZone -Name $reverseZoneName -ErrorAction SilentlyContinue)) {
    Add-DnsServerPrimaryZone -NetworkId "192.168.250.0/24" -ZoneFile "$reverseZoneName.dns" -DynamicUpdate Secure
} else {
    Set-DnsServerPrimaryZone -Name $reverseZoneName -DynamicUpdate Secure
}

# DNS-Weiterleitungen konfigurieren
Set-DnsServerForwarder -IPAddress $forwarders

# Beispiel-Hosteintrag hinzufügen oder aktualisieren
$hostName = "server1"
$hostIP = "192.168.250.203"
if (-not (Get-DnsServerResourceRecord -ZoneName $zoneName -Name $hostName -ErrorAction SilentlyContinue)) {
    Add-DnsServerResourceRecordA -ZoneName $zoneName -Name $hostName -IPv4Address $hostIP
} else {
    Set-DnsServerResourceRecord -ZoneName $zoneName -OldInputObject (Get-DnsServerResourceRecord -ZoneName $zoneName -Name $hostName) -NewInputObject (New-DnsServerResourceRecordA -Name $hostName -IPv4Address $hostIP)
}

Write-Output "DNS-Konfiguration abgeschlossen."
