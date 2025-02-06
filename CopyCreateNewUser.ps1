# Copy-ADUserTemplate.ps1
#
# .SYNOPSIS
# Erstellt neue Active Directory Benutzer basierend auf einem Vorlagenbenutzer.
#
# .DESCRIPTION
# Dieses Skript erstellt einen oder mehrere neue Active Directory Benutzer durch Kopieren eines vorhandenen Vorlagenbenutzers.
# Es unterstützt sowohl interaktive als auch Batch-Verarbeitung via CSV und kopiert alle relevanten AD-Attribute und Gruppenmitgliedschaften.
#
# .PARAMETER TemplateUser
# SAMAccountName des Vorlagenbenutzers
#
# .PARAMETER NewUser
# SAMAccountName des neuen Benutzers
#
# .PARAMETER Password
# Passwort für den neuen Benutzer
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
# .\Copy-ADUserTemplate.ps1 -TemplateUser "muster" -NewUser "neueruser" -Password "Sicheres!Pw123"
# Erstellt einen neuen Benutzer basierend auf dem Vorlagenbenutzer
#
# .EXAMPLE
# .\Copy-ADUserTemplate.ps1 -CsvPath ".\neue_benutzer.csv"
# Erstellt mehrere Benutzer aus CSV-Datei
#
# .NOTES
# Version: 2.0
# Autor: Bolt
# Benötigt: Active Directory PowerShell Modul
# Kompatibel mit PowerShell 5.1
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false, ParameterSetName = 'Single')]
    [string]$TemplateUser,
    
    [Parameter(Mandatory = $false, ParameterSetName = 'Single')]
    [string]$NewUser,
    
    [Parameter(Mandatory = $false, ParameterSetName = 'Single')]
    [string]$Password,
    
    [Parameter(Mandatory = $false, ParameterSetName = 'CSV')]
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
    param([string]$SamAccountName)
    try {
        Write-Log "Rufe Vorlagenbenutzer ab: $SamAccountName"
        $user = Get-ADUser -Identity $SamAccountName -Properties *
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
            SamAccountName = $NewUserProps.sAMAccountName
            UserPrincipalName = "$($NewUserProps.sAMAccountName)@$((Get-ADDomain).DNSRoot)"
        }
        
        # Benutzer erstellen
        $newUser = New-ADUser @userParams -PassThru
        
        # Passwort setzen
        $securePassword = ConvertTo-SecureString $NewUserProps.Password -AsPlainText -Force
        Set-ADAccountPassword -Identity $newUser -NewPassword $securePassword
        
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
    
    if ($PSCmdlet.ParameterSetName -eq 'Single') {
        # Einzelbenutzer-Modus
        if (-not $TemplateUser -or -not $NewUser -or -not $Password) {
            Write-Error "Für Einzelbenutzer-Modus werden TemplateUser, NewUser und Password benötigt"
            exit 1
        }
        
        # Vorlagenbenutzer laden
        $templateUserObj = Get-TemplateUser -SamAccountName $TemplateUser
        Write-Log "Vorlagenbenutzer gefunden: $TemplateUser"
        
        # Passwort-Komplexität prüfen
        if (-not (Test-PasswordComplexity $Password)) {
            Write-Error "Passwort entspricht nicht den Anforderungen"
            exit 1
        }
        
        # Neuen Benutzer erstellen
        $newUserProps = @{
            sAMAccountName = $NewUser
            Password = $Password
        }
        
        $newUser = New-ClonedUser -TemplateUser $templateUserObj -NewUserProps $newUserProps -TargetOU $TargetOU
        Copy-UserGroups -SourceUser $templateUserObj -TargetUser $newUser
    }
    elseif ($PSCmdlet.ParameterSetName -eq 'CSV') {
        # CSV-Modus
        if (-not $CsvPath) {
            Write-Error "Für CSV-Modus wird CsvPath benötigt"
            exit 1
        }
        
        Write-Log "Verarbeite CSV-Datei: $CsvPath"
        $users = Import-Csv -Path $CsvPath -Delimiter ';' -Encoding UTF8
        
        foreach ($user in $users) {
            $templateUserObj = Get-TemplateUser -SamAccountName $user.TemplateUser
            
            if (-not (Test-PasswordComplexity $user.Password)) {
                Write-Log "Passwort entspricht nicht den Anforderungen für: $($user.NewUser)" -Level "FEHLER"
                continue
            }
            
            $newUserProps = @{
                sAMAccountName = $user.NewUser
                Password = $user.Password
            }
            
            $newUser = New-ClonedUser -TemplateUser $templateUserObj -NewUserProps $newUserProps -TargetOU $TargetOU
            Copy-UserGroups -SourceUser $templateUserObj -TargetUser $newUser
        }
    }
    
    Write-Log "Skript erfolgreich beendet" -Level "ERFOLG"
}
catch {
    Write-Log "Fataler Fehler: $_" -Level "FEHLER"
    throw
}
