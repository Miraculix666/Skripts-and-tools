# Definieren Sie den Namen des Remote-Computers und der Software
$computerName = "RemoteComputerName"
$softwareName = "SoftwareName"

# Erstellen Sie eine Remote-Registry-Verbindung
$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $computerName)

# Durchsuchen Sie die Uninstall-Schlüssel in der Registry
$regKey = $reg.OpenSubKey('SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall')

# Durchsuchen Sie alle Uninstall-Schlüssel
$softwareFound = $false
$regKey.GetSubKeyNames() | ForEach-Object {
    $key = $regKey.OpenSubKey($_)
    $displayName = $key.GetValue('DisplayName')
    if ($displayName -and $displayName -like "*$softwareName*") {
        Write-Output "Die Software $softwareName ist auf $computerName installiert."
        $softwareFound = $true
    }
}

# Überprüfen Sie, ob die Software gefunden wurde
if (-not $softwareFound) {
    Write-Output "Die Software $softwareName ist auf $computerName nicht installiert."
}


# Definieren Sie den Namen des Remote-Computers und der Software
$computerName = "RemoteComputerName"
$softwareName = "SoftwareName"

# Erstellen Sie eine Remote-Registry-Verbindung
$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $computerName)

# Durchsuchen Sie die Uninstall-Schlüssel in der Registry
$regKey = $reg.OpenSubKey('SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall')

# Durchsuchen Sie alle Uninstall-Schlüssel
$softwareFound = $false
$regKey.GetSubKeyNames() | ForEach-Object {
    $key = $regKey.OpenSubKey($_)
    $displayName = $key.GetValue('DisplayName')
    if ($displayName -and $displayName -like "*$softwareName*") {
        Write-Output "Die Software $softwareName ist auf $computerName installiert."
        $softwareFound = $true
    }
}

# Überprüfen Sie, ob die Software gefunden wurde
if (-not $softwareFound) {
    Write-Output "Die Software $softwareName ist auf $computerName nicht installiert."
}


