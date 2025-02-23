# Active Directory Benutzerverwaltungsskript
# Autor: [Ihr Name]
# Version: 1.1
# Beschreibung:
# - Exportiert bestehende AD-Benutzer in eine CSV-Datei basierend auf einem Template-Benutzer.
# - Erstellt neue Benutzer interaktiv, mit Parametern oder per CSV-Datei.
# - Weist Gruppen basierend auf dem Template-Benutzer zu.
# - Unterstützt Logging, Fehlermeldungen und ausführliche Ausgabe (Verbose).
# - Lokalisierte CSV-Datei und Ausgabe (Deutsch).

[CmdletBinding()]
param (
    [string]$TemplateUser,
    [string]$ExportPath = "ADUsersExport.csv",
    [string]$CsvPath,
    [switch]$Verbose
)

# Prüfen, ob das Active Directory Modul verfügbar ist
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host "Das Active Directory Modul ist nicht installiert. Bitte installieren Sie es zuerst." -ForegroundColor Red
    exit 1
}

# Modul laden
Import-Module ActiveDirectory

# Funktion zur interaktiven Abfrage fehlender Parameter
function Get-ParameterIfMissing {
    param (
        [string]$ParameterName,
        [string]$PromptMessage
    )
    if (-not $PSBoundParameters.ContainsKey($ParameterName)) {
        return Read-Host -Prompt $PromptMessage
    }
    return (Get-Variable -Name $ParameterName -ValueOnly)
}

# Logging-Funktion
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$Timestamp [$Level] $Message" | Out-File -Append -FilePath "ADUserManagement.log"
    if ($Verbose) { Write-Host "$Timestamp [$Level] $Message" -ForegroundColor Yellow }
}

# Funktion zum Exportieren von AD-Benutzern basierend auf einem Template-User
function Export-ADUsers {
    $TemplateUser = Get-ParameterIfMissing -ParameterName "TemplateUser" -PromptMessage "Geben Sie den Template-Benutzer an"
    Write-Log "Starte Benutzerexport basierend auf Template: $TemplateUser"
    try {
        $Template = Get-ADUser -Identity $TemplateUser -Properties MemberOf
        if (-not $Template) { throw "Template-Benutzer nicht gefunden!" }
        
        Get-ADUser -Filter * -Properties SamAccountName, UserPrincipalName, Name, MemberOf |
        Where-Object { $_.MemberOf -contains $Template.MemberOf } |
        Select-Object SamAccountName, UserPrincipalName, Name, @{Name='Groups'; Expression={$_.MemberOf -join ';'}} |
        Export-Csv -Path $ExportPath -Delimiter ";" -NoTypeInformation -Encoding UTF8
        
        Write-Log "Benutzerdaten erfolgreich exportiert: $ExportPath"
    } catch {
        Write-Log "Fehler beim Export: $_" -Level "ERROR"
    }
}

# Funktion zur Erstellung neuer Benutzer mit Gruppen des Template-Benutzers
function Create-ADUser {
    param (
        [string]$SamAccountName,
        [string]$UserPrincipalName,
        [string]$Name,
        [string]$OU,
        [string[]]$Groups,
        [string]$Password = "Passwort123!"
    )
    Write-Log "Erstelle Benutzer: $Name ($SamAccountName)"
    try {
        New-ADUser -Name $Name `
                   -SamAccountName $SamAccountName `
                   -UserPrincipalName $UserPrincipalName `
                   -Path $OU `
                   -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -Force) `
                   -Enabled $true

        foreach ($group in $Groups) {
            if (Get-ADGroup -Filter { Name -eq $group }) {
                Add-ADGroupMember -Identity $group -Members $SamAccountName
            } else {
                Write-Log "Gruppe nicht gefunden: $group" -Level "WARNING"
            }
        }

        Write-Log "Benutzer $Name erfolgreich erstellt"
    } catch {
        Write-Log "Fehler beim Erstellen von $Name: $_" -Level "ERROR"
    }
}

# Funktion zur Erstellung mehrerer Benutzer aus CSV
function Create-ADUsersFromCSV {
    $CsvPath = Get-ParameterIfMissing -ParameterName "CsvPath" -PromptMessage "Geben Sie den Pfad zur CSV-Datei an"
    Write-Log "Erstelle Benutzer aus CSV: $CsvPath"
    if (-Not (Test-Path $CsvPath)) {
        Write-Log "Fehler: Datei $CsvPath nicht gefunden!" -Level "ERROR"
        return
    }
    
    $users = Import-Csv -Path $CsvPath -Delimiter ";"
    foreach ($user in $users) {
        Create-ADUser -SamAccountName $user.SamAccountName `
                      -UserPrincipalName $user.UserPrincipalName `
                      -Name $user.Name `
                      -OU $user.OU `
                      -Groups ($user.Groups -split ';') `
                      -Password $user.Password
    }
}

# Hilfe-Funktion
function Show-Help {
    Write-Host "AD Benutzerverwaltungsskript - Hilfe"
    Write-Host "==================================="
    Write-Host "Verwendung:"
    Write-Host "  - Template-User exportieren: .\Script.ps1 -TemplateUser 'Mustermann'"
    Write-Host "  - Benutzer aus CSV erstellen: .\Script.ps1 -CsvPath 'Users.csv'"
    Write-Host "  - Interaktiver Modus: .\Script.ps1"
}

# Hauptprogramm
if ($TemplateUser) {
    Export-ADUsers
}
if ($CsvPath) {
    Create-ADUsersFromCSV
}
if (-not $TemplateUser -and -not $CsvPath) {
    Show-Help
}
