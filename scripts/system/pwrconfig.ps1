# Aktiven Energiesparplan ermitteln
$activeScheme = (powercfg /getactivescheme) -match '{.*}' | Out-Null
$activeScheme = $matches[0]

# Funktion zum Setzen der Werte
function Set-PowerSetting {
    param (
        [string]$SchemeGuid,
        [string]$SubGroupGuid,
        [string]$SettingGuid,
        [int]$ACValue,
        [int]$DCValue
    )
    powercfg /setacvalueindex $SchemeGuid $SubGroupGuid $SettingGuid $ACValue
    powercfg /setdcvalueindex $SchemeGuid $SubGroupGuid $SettingGuid $DCValue
}

# Verhalten beim Zuklappen (0 = keine Aktion)
Set-PowerSetting -SchemeGuid $activeScheme `
    -SubGroupGuid "4f971e89-eebd-4455-a8de-9e59040e7347" `
    -SettingGuid "5ca83367-6e45-459f-a27b-476b1d01c936" `
    -ACValue 0 -DCValue 0

# Standby-Button (1 = Standby)
Set-PowerSetting -SchemeGuid $activeScheme `
    -SubGroupGuid "4f971e89-eebd-4455-a8de-9e59040e7347" `
    -SettingGuid "96996bc0-ad50-47ec-923b-6f41874dd9eb" `
    -ACValue 1 -DCValue 1

# Power-Button (3 = Ruhezustand)
Set-PowerSetting -SchemeGuid $activeScheme `
    -SubGroupGuid "4f971e89-eebd-4455-a8de-9e59040e7347" `
    -SettingGuid "7648efa3-dd9c-4e3e-b566-50f929386280" `
    -ACValue 3 -DCValue 3

# Bildschirm ausschalten nach (AC: 10 min, DC: 5 min)
Set-PowerSetting -SchemeGuid $activeScheme `
    -SubGroupGuid "7516b95f-f776-4464-8c53-06167f40cc99" `
    -SettingGuid "3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e" `
    -ACValue 600 -DCValue 300

# Standby nach (AC: 30 min, DC: 15 min)
Set-PowerSetting -SchemeGuid $activeScheme `
    -SubGroupGuid "238c9fa8-0aad-41ed-83f4-97be242c8f20" `
    -SettingGuid "29f6c1db-86da-48c5-9fdb-f2b67b1f44da" `
    -ACValue 1800 -DCValue 900

# Ruhezustand nach (AC: 60 min, DC: 30 min)
Set-PowerSetting -SchemeGuid $activeScheme `
    -SubGroupGuid "238c9fa8-0aad-41ed-83f4-97be242c8f20" `
    -SettingGuid "9d7815a6-7ee4-497e-8888-515a05f02364" `
    -ACValue 3600 -DCValue 1800

# Änderungen aktivieren
powercfg /setactive $activeScheme
