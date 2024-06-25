# Definieren Sie den Namen des Remote-Computers und der Software
$computerName = "RemoteComputerName"
$softwareName = "SoftwareName"

# Test der Remote-Registry-Verbindung 
try {
    $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $computerName)
    Write-Output "Die Remote-Verbindung zur Registry von $computerName wurde erfolgreich hergestellt."
} catch {
    Write-Output "Die Remote-Verbindung zur Registry von $computerName konnte nicht hergestellt werden. Fehler: $_"
}
Write-Output "Die Software $softwareName ist auf $computerName nicht installiert."
}

# Erstellen Sie eine Remote-Registry-Verbindung
$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $computerName)

# Öffnen Sie den Uninstall-Schlüssel in der Registry
$regKey = $reg.OpenSubKey('SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall')

# Durchsuchen Sie alle Uninstall-Schlüssel
$softwareFound = $false
$regKey.GetSubKeyNames() | ForEach-Object {
    $key = $regKey.OpenSubKey($_)
    $displayName = $key.GetValue('DisplayName')
    $displayVersion = $key.GetValue('DisplayVersion')
    if ($displayName -and $displayName -like "*$softwareName*") {
        Write-Output "Die Software $displayName Version $displayVersion ist auf $computerName installiert."
        $softwareFound = $true
    }
}

# Überprüfen Sie, ob die Software gefunden wurde
if (-not $softwareFound) {
    Write-Output "Es wurde keine Software gefunden, die den String $softwareName auf $computerName enthält."
}



##### Abfrage mit WMI

# Versuchen Sie, eine Remote-Registry-Verbindung herzustellen
try {
    $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $computerName)
    Write-Output "Die Remote-Verbindung zur Registry von $computerName wurde erfolgreich hergestellt."
} catch {
    Write-Output "Die Remote-Verbindung zur Registry von $computerName konnte nicht hergestellt werden. Fehler: $_"
}


# Erstellen Sie eine WMI-Abfrage
$query = "SELECT * FROM Win32_Product WHERE (Name LIKE '%$softwareName%')"

# Führen Sie die Abfrage auf dem Remote-Computer aus
$software = Get-WmiObject -Query $query -ComputerName $computerName

# Überprüfen Sie, ob die Software gefunden wurde
if ($software) {
    Write-Output "Die Software $softwareName ist auf $computerName installiert."
} else {
    Write-Output "Die Software $softwareName ist auf $computerName nicht installiert."
}

