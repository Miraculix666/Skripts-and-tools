# DNS-Serverrolle installieren, falls nicht vorhanden
Install-WindowsFeature -Name DNS -IncludeManagementTools

# DNS-Serverdienst starten
Start-Service -Name DNS

# Zone neu laden
dnscmd /ZoneReload lafp.schul.polizei.local

# Zone-Eigenschaften setzen
dnscmd /ZoneResetType lafp.schul.polizei.local /DsPrimary

# DNS-Weiterleitungen konfigurieren
$forwarders = @("1.1.1.1", "8.8.8.8") # Cloudflare und Google DNS
Set-DnsServerForwarder -IPAddress $forwarders

# Beispiel-Hosteintrag hinzufügen oder aktualisieren
$zoneName = "lafp.schul.polizei.local"
$hostName = "SDOM"
$hostIP = "192.168.250.203"
if (-not (Get-DnsServerResourceRecord -ZoneName $zoneName -Name $hostName -ErrorAction SilentlyContinue)) {
    Add-DnsServerResourceRecordA -ZoneName $zoneName -Name $hostName -IPv4Address $hostIP
} else {
    $record = Get-DnsServerResourceRecord -ZoneName $zoneName -Name $hostName
    Remove-DnsServerResourceRecord -ZoneName $zoneName -InputObject $record
    Add-DnsServerResourceRecordA -ZoneName $zoneName -Name $hostName -IPv4Address $hostIP
}

Write-Output "DNS-Konfiguration abgeschlossen."
