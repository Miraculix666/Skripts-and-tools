# Definieren Sie den Namen des Remote-Computers und der Software
$computerName = "RemoteComputerName"
$softwareName = "SoftwareName"

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
