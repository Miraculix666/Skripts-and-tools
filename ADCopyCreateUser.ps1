# Active Directory Benutzerverwaltungsskript
# Autor: [Ihr Name]
# Beschreibung:
# - Exportiert bestehende AD-Benutzer in eine CSV-Datei basierend auf einem Template-Benutzer
# - Erstellt neue Benutzer interaktiv, mit Parametern oder per CSV-Datei
# - Weist Gruppen basierend auf dem Template-Benutzer zu
# - Unterstützt Logging, Fehlermeldungen und ausführliche Ausgabe (Verbose)

param (
    [string]$TemplateUser,
    [string]$ExportPath = "ADUsersExport.csv",
    [string]$CsvPath
)

Import-Module ActiveDirectory

# Funktion zur interaktiven Abfrage fehlender Parameter
function Get-ParameterIfMissing {
    param (
        [string]$ParameterName,
        [string]$PromptMessage
    )
    if (-not (Get-Variable -Name $ParameterName -ValueOnly -ErrorAction SilentlyContinue)) {
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
    Write-Host "$Timestamp [$Level] $Message" -ForegroundColor Yellow
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
            Add-ADGroupMember -Identity $group -Members $SamAccountName
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

# Hauptprogramm
if (-not $TemplateUser) {
    $TemplateUser = Get-ParameterIfMissing -ParameterName "TemplateUser" -PromptMessage "Geben Sie den Template-Benutzer an"
}
if (-not $ExportPath) {
    $ExportPath = Get-ParameterIfMissing -ParameterName "ExportPath" -PromptMessage "Geben Sie den Exportpfad an"
}
if (-not $CsvPath) {
    $CsvPath = Get-ParameterIfMissing -ParameterName "CsvPath" -PromptMessage "Geben Sie den Pfad zur CSV-Datei an"
}

if ($TemplateUser) {
    Export-ADUsers
}
if ($CsvPath) {
    Create-ADUsersFromCSV
}
