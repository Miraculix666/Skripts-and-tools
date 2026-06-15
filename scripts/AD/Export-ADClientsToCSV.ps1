# Input
$searchString = Read-Host "Enter a partial client name"

# Get AD computers matching the search string
$computers = Get-ADComputer -Filter "Name -like '*$searchString*'" -Properties Name, Description, IPv4Address

# Create CSV file
$csvPath = ".\ADComputers.csv"
"Client-Name;Beschreibung;IP-Adresse;Mac-Adresse" | Out-File $csvPath -Encoding UTF8

# Function to get MAC address
function Get-MacAddress {
    param ($ComputerName, $IPAddress)
    
    # Method 1: Use Get-NetNeighbor (fastest, but requires admin rights and works only for local subnet)
    if ($IPAddress) {
        try {
            $mac = Get-NetNeighbor -IPAddress $IPAddress -ErrorAction Stop | Select-Object -ExpandProperty LinkLayerAddress
            if ($mac) { return $mac }
        } catch {}
    }
    
    # Method 2: Use WMI (works remotely but can be slow)
    try {
        $mac = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $ComputerName -Filter "IPEnabled='True'" -ErrorAction Stop |
               Select-Object -First 1 -ExpandProperty MACAddress
        if ($mac) { return $mac }
    } catch {}
    
    # Method 3: Use PowerShell remoting (requires remoting to be enabled)
    try {
        $mac = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            (Get-NetAdapter | Where-Object Status -eq 'Up' | Select-Object -First 1).MacAddress
        } -ErrorAction Stop
        if ($mac) { return $mac }
    } catch {}
    
    return "Not Available"
}

# Process each computer
foreach ($computer in $computers) {
    $macAddress = Get-MacAddress -ComputerName $computer.Name -IPAddress $computer.IPv4Address
    
    # Prepare CSV line
    $csvLine = "{0};{1};{2};{3}" -f $computer.Name, $computer.Description, $computer.IPv4Address, $macAddress
    
    # Append to CSV file
    $csvLine | Out-File $csvPath -Append -Encoding UTF8
}

Write-Host "CSV file created at $csvPath"
