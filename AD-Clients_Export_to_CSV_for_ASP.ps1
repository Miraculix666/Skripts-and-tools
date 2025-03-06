# Prompt for input
$searchString = Read-Host "Enter a partial client name"

# Get AD computers matching the search string
$computers = Get-ADComputer -Filter "Name -like '*$searchString*'" -Properties Description, IPv4Address

# Create CSV file
$csvPath = ".\ADComputers.csv"
"Client-Name;Beschreibung;IP-Adresse;Mac-Adresse" | Out-File $csvPath -Encoding UTF8

# Process each computer
foreach ($computer in $computers) {
    $macAddress = ""
    try {
        $macAddress = (Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $computer.Name -Filter "IPEnabled='True'" | Select-Object -First 1).MACAddress
    } catch {
        # If unable to retrieve MAC address, leave it blank
    }

    # Prepare CSV line
    $csvLine = "{0};{1};{2};{3}" -f $computer.Name, $computer.Description, $computer.IPv4Address, $macAddress

    # Append to CSV file
    $csvLine | Out-File $csvPath -Append -Encoding UTF8
}

Write-Host "CSV file created at $csvPath"
