# Copy-ADUserTemplate.ps1
#
# .SYNOPSIS
# Erstellt neue Active Directory Benutzer basierend auf einem Vorlagenbenutzer.
#
# .DESCRIPTION
# Dieses Skript erstellt einen oder mehrere neue Active Directory Benutzer durch Kopieren eines vorhandenen Vorlagenbenutzers.
# Es unterstützt sowohl interaktive als auch Batch-Verarbeitung via CSV und kopiert alle relevanten AD-Attribute und Gruppenmitgliedschaften.
#
# .PARAMETER TemplateUserDN
# Distinguished Name des Vorlagenbenutzers
#
# .PARAMETER CsvPath
# (Optional) Pfad zur CSV-Datei mit neuen Benutzerdaten (Semikolon als Trennzeichen)
#
# .PARAMETER TargetOU
# (Optional) Distinguished Name der Ziel-OU für neue Benutzer
#
# .PARAMETER LogPath
# Pfad für die Logdatei. Standard: ".\ADUserClone_<Zeitstempel>.log"
#
# .EXAMPLE
# .\Copy-ADUserTemplate.ps1 -TemplateUserDN "CN=Musterbenutzer,OU=Benutzer,DC=firma,DC=local"
# Erstellt neue Benutzer interaktiv basierend auf der angegebenen Vorlage
#
# .EXAMPLE
# .\Copy-ADUserTemplate.ps1 -TemplateUserDN "CN=Musterbenutzer,OU=Benutzer,DC=firma,DC=local" -CsvPath ".\neue_benutzer.csv"
# Erstellt mehrere Benutzer aus CSV-Datei mit der angegebenen Vorlage
#
# .NOTES
# Version: 2.0
# Autor: Bolt
# Benötigt: Active Directory PowerShell Modul
# Kompatibel mit PowerShell 5.1
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [string]$TemplateUserDN,
    
    [Parameter(Mandatory = $false)]
    [string]$CsvPath,
    
    [Parameter(Mandatory = $false)]
    [string]$TargetOU,
    
    [Parameter(Mandatory = $false)]
    [string]$LogPath = ".\ADUserClone_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
)

# Setze deutsche Kultureinstellungen
$culture = [System.Globalization.CultureInfo]::GetCultureInfo("de-DE")
[System.Threading.Thread]::CurrentThread.CurrentCulture = $culture
[System.Threading.Thread]::CurrentThread.CurrentUICulture = $culture

# Import des AD-Moduls
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Verbose "Active Directory Modul erfolgreich geladen."
}
catch {
    Write-Error "Active Directory Modul konnte nicht geladen werden: $_"
    exit 1
}

# Logging-Funktion
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARNUNG", "FEHLER", "ERFOLG")]
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp][$Level] $Message"
    Write-Verbose $logMessage
    Add-Content -Path $LogPath -Value $logMessage -Encoding UTF8
}

# Pflichtattribute für neue Benutzer
$mandatoryAttributes = @(
    'sAMAccountName',
    'userPrincipalName',
    'givenName',
    'sn',
    'displayName'
)

# Optionale Attribute zum Kopieren
$optionalAttributes = @(
    'description',
    'mail',
    'physicalDeliveryOfficeName',
    'department',
    'company',
    'title',
    'streetAddress',
    'l',
    'postalCode',
    'c',
    'telephoneNumber',
    'mobile',
    'homePhone',
    'wWWHomePage',
    'employeeID'
)

# Passwort-Komplexitätsprüfung
function Test-PasswordComplexity {
    param([string]$Password)
    
    $rules = @(
        { $Password.Length -ge 12 }                # Mindestlänge 12 Zeichen
        { $Password -match '[A-Z]' }               # Großbuchstabe
        { $Password -match '[a-z]' }               # Kleinbuchstabe
        { $Password -match '\d' }                  # Ziffer
        { $Password -match '[\W_]' }               # Sonderzeichen
    )
    
    $valid = $true
    foreach ($rule in $rules) {
        if (-not (& $rule)) {
            $valid = $false
            break
        }
    }
    
    Write-Verbose "Passwort-Komplexität: $(if($valid){'OK'}else{'Nicht ausreichend'})"
    return $valid
}

# Vorlagenbenutzer abrufen
function Get-TemplateUser {
    param([string]$DN)
    try {
        Write-Log "Rufe Vorlagenbenutzer ab: $DN"
        $user = Get-ADUser -Identity $DN -Properties *
        if (-not $user) {
            throw "Vorlagenbenutzer nicht gefunden"
        }
        return $user
    }
    catch {
        Write-Log "Fehler beim Abrufen des Vorlagenbenutzers: $_" -Level "FEHLER"
        throw
    }
}

# Gruppenmitgliedschaften kopieren
function Copy-UserGroups {
    param(
        $SourceUser,
        $TargetUser
    )
    try {
        Write-Log "Kopiere Gruppenmitgliedschaften"
        $groups = Get-ADUser $SourceUser -Properties MemberOf | 
            Select-Object -ExpandProperty MemberOf
        
        foreach ($group in $groups) {
            Add-ADGroupMember -Identity $group -Members $TargetUser
            Write-Log "Gruppe hinzugefügt: $group" -Level "INFO"
        }
    }
    catch {
        Write-Log "Fehler beim Kopieren der Gruppenmitgliedschaften: $_" -Level "FEHLER"
        throw
    }
}

# Neuen Benutzer erstellen
function New-ClonedUser {
    param(
        $TemplateUser,
        $NewUserProps,
        $TargetOU
    )
    
    try {
        Write-Log "Erstelle neuen Benutzer: $($NewUserProps.sAMAccountName)"
        
        # Basisparameter
        $userParams = @{
            Instance = $TemplateUser
            Path = if ($TargetOU) { $TargetOU } else { ($TemplateUser.DistinguishedName -split ',', 2)[1] }
            Enabled = $true
            ChangePasswordAtLogon = $true
        }
        
        # Pflichtattribute
        foreach ($attr in $mandatoryAttributes) {
            if ($NewUserProps.$attr) {
                $userParams[$attr] = $NewUserProps.$attr
            }
        }
        
        # Optionale Attribute
        foreach ($attr in $optionalAttributes) {
            if ($NewUserProps.$attr) {
                $userParams[$attr] = $NewUserProps.$attr
            }
        }
        
        # Benutzer erstellen
        $newUser = New-ADUser @userParams -PassThru
        
        # Passwort setzen
        $password = if ($NewUserProps.Password) {
            ConvertTo-SecureString $NewUserProps.Password -AsPlainText -Force
        }
        else {
            Read-Host "Passwort für $($NewUserProps.sAMAccountName)" -AsSecureString
        }
        
        Set-ADAccountPassword -Identity $newUser -NewPassword $password
        
        Write-Log "Benutzer erfolgreich erstellt: $($NewUserProps.sAMAccountName)" -Level "ERFOLG"
        return $newUser
    }
    catch {
        Write-Log "Fehler beim Erstellen des Benutzers: $_" -Level "FEHLER"
        throw
    }
}

# Hauptausführung
try {
    Write-Log "Skript gestartet"
    
    # Vorlagenbenutzer laden
    $templateUser = Get-TemplateUser -DN $TemplateUserDN
    Write-Log "Vorlagenbenutzer gefunden: $($templateUser.sAMAccountName)"
    
    if ($CsvPath) {
        # CSV-Modus
        Write-Log "Verarbeite CSV-Datei: $CsvPath"
        $users = Import-Csv -Path $CsvPath -Delimiter ';' -Encoding UTF8
        
        foreach ($user in $users) {
            # Pflichtattribute prüfen
            $missingAttrs = $mandatoryAttributes | Where-Object { -not $user.$_ }
            if ($missingAttrs) {
                Write-Log "Fehlende Pflichtattribute für Benutzer: $($missingAttrs -join ', ')" -Level "FEHLER"
                continue
            }
            
            # Passwort-Komplexität prüfen
            if ($user.Password -and -not (Test-PasswordComplexity $user.Password)) {
                Write-Log "Passwort entspricht nicht den Anforderungen für: $($user.sAMAccountName)" -Level "FEHLER"
                continue
            }
            
            $newUser = New-ClonedUser -TemplateUser $templateUser -NewUserProps $user -TargetOU $TargetOU
            Copy-UserGroups -SourceUser $templateUser -TargetUser $newUser
        }
    }
    else {
        # Interaktiver Modus
        Write-Log "Starte interaktiven Modus"
        
        $newUserProps = @{}
        
        # Pflichtattribute abfragen
        foreach ($attr in $mandatoryAttributes) {
            $newUserProps[$attr] = Read-Host "Bitte $attr eingeben"
        }
        
        # Optionale Attribute mit Vorlagewerten
        foreach ($attr in $optionalAttributes) {
            $default = $templateUser.$attr
            $input = Read-Host "Bitte $attr eingeben (Standard: $default)"
            $newUserProps[$attr] = if ($input) { $input } else { $default }
        }
        
        $newUser = New-ClonedUser -TemplateUser $templateUser -NewUserProps $newUserProps -TargetOU $TargetOU
        Copy-UserGroups -SourceUser $templateUser -TargetUser $newUser
    }
    
    Write-Log "Skript erfolgreich beendet" -Level "ERFOLG"
}
catch {
    Write-Log "Fataler Fehler: $_" -Level "FEHLER"
    throw
}
