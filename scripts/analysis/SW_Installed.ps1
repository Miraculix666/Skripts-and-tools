# Dieses Skript überprüft, ob eine bestimmte Software auf einem oder mehreren Computern installiert ist.
# Es nimmt den Computernamen und den Softwarenamen als Kommandozeilenoptionen an.
# Alternativ kann eine CSV-Datei anstelle des Computernamens angegeben werden.
# Die CSV-Datei sollte das Format "Client-Name;Beschreibung;IP-Adresse;Mac-Adresse" haben und eine Kopfzeile enthalten.

# Definieren Sie die Eingabeparameter
param (
    [Parameter(Mandatory=$true)]
    [string]$computerName,
    [Parameter(Mandatory=$true)]
    [string]$softwareName,
    [Parameter(Mandatory=$false)]
    [string]$csvFile
)

# Funktion zum Überprüfen der installierten Software auf einem Computer
function Check-InstalledSoftware {
    param (
        [Parameter(Mandatory=$true)]
        [string]$computerName,
        [Parameter(Mandatory=$true)]
        [string]$softwareName
    )

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
            $softwareFound = $true
            # Erstellen Sie ein benutzerdefiniertes PowerShell-Objekt für die tabellarische Ausgabe
            [PSCustomObject]@{
                'ComputerName' = $computerName
                'SoftwareName' = $displayName
                'Version' = $displayVersion
            }
        }
    }

    # Überprüfen Sie, ob die Software gefunden wurde
    if (-not $softwareFound) {
        Write-Output "Es wurde keine Software gefunden, die den String $softwareName auf $computerName enthält."
    }
}

# Überprüfen Sie, ob eine CSV-Datei angegeben wurde
if ($csvFile) {
    # Importieren Sie die CSV-Datei
    $computers = Import-Csv -Path $csvFile -Delimiter ';'

    # Überprüfen Sie die installierte Software auf jedem Computer in der CSV-Datei
    $computers | ForEach-Object {
        Check-InstalledSoftware -computerName $_.'Client-Name' -softwareName $softwareName
    }
} else {
    # Überprüfen Sie die installierte Software auf dem angegebenen Computer
    Check-InstalledSoftware -computerName $computerName -softwareName $softwareName
}
