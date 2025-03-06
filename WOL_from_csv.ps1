# Function to send WOL packet
function Send-WOLPacket {
    param(
        [string]$MacAddress
    )
    
    $broadcast = [System.Net.IPAddress]::Broadcast
    $mac = $MacAddress -replace "[:-]", ""
    $target = [byte[]]($mac -split "([a-f0-9]{2})" | Where-Object { $_ -ne "" } | ForEach-Object { [Convert]::ToByte($_, 16) })
    $packet = [byte[]](@(0xFF) * 6) + ($target * 16)
    
    $UdpClient = New-Object System.Net.Sockets.UdpClient
    $UdpClient.Connect($broadcast, 9)
    $UdpClient.Send($packet, $packet.Length) | Out-Null
    $UdpClient.Close()
}

# Read CSV file
$csvPath = "path\to\your\csv\file.csv"
$computers = Import-Csv $csvPath

# Send WOL packet for each computer
foreach ($computer in $computers) {
    Write-Host "Sending WOL packet to $($computer.'Client-Name')"
    Send-WOLPacket -MacAddress $computer.'Mac-Adresse'
}
