<#
.SYNOPSIS
    Erstellt neue AD-Benutzer als exakte Kopien eines Vorlagenbenutzers.

.DESCRIPTION
    Dieses Skript erstellt einen oder mehrere neue Active Directory Benutzer, indem es einen vorhandenen Vorlagenbenutzer kopiert.
    Es übernimmt alle kopierbaren Attribute, z. B. Gruppenmitgliedschaften, OU‑Platzierung, AD‑Rechte und optional weitere 
    benutzerdefinierte Werte. Dabei werden zwingend neue Attribute (SAMAccountName, AccountPassword, Name) von optionalen 
    Attributen (GivenName, Surname, DisplayName, Abteilung, Office, etc.) unterschieden.
    
    Das Skript unterstützt drei Aufrufvarianten:
      • Keine Parameter: Interaktiver Modus – es werden alle fehlenden Parameter abgefragt.
      • Einzelbenutzer: Alle nötigen Parameter (TemplateUser, NewUserName, NewUserPassword) werden übergeben.
      • CSV-Batch: Mit dem Parameter –CsvPath (Die CSV-Datei verwendet in der deutschen Version den Semikolon‑Delimiter).
        
    WICHTIG: Im CSV‑Modus müssen folgende Spalten vorhanden sein (mit Semikolon als Trenner):
        TemplateUser;NewUserName;NewUserPassword;GivenName;Surname;Department;Office;WeitereAttribute...
        
    Das Skript wandelt Plaintext ‑ Kennwörter in SecureStrings um und gibt detaillierte Log‑ und “verbose” Ausgaben aus.

.PARAMETER CsvPath
    (Optional) Pfad zu einer CSV-Datei, in der die neuen Benutzer definiert sind. 
    Die CSV muss die Pflichtspalten TemplateUser, NewUserName und NewUserPassword enthalten,
    optional werden weitere Spalten wie GivenName, Surname, Department, Office etc. unterstützt.

.PARAMETER TemplateUser
    (Optional, außer CSV‑Modus) Der SAMAccountName des Vorlagenbenutzers.

.PARAMETER NewUserName
    (Optional, außer CSV‑Modus) Der SAMAccountName des neuen AD-Benutzers.

.PARAMETER NewUserPassword
    (Optional, außer CSV‑Modus) Das Plaintext‑Kennwort für den neuen Benutzer.
    Das Skript konvertiert es in einen SecureString. (Fehler, "could not convert system.string to system.security.securestring" 
    werden vermieden.)

.PARAMETER Verify
    Schalter zur Aktivierung einer abschließenden Verifikation der erstellten Benutzer (z. B. Attributvergleich und Gruppen).

.EXAMPLE
    # Interaktiver Modus:
    PS C:\> .\CopyCreateNewUser.ps1

.EXAMPLE
    # Einzelbenutzermodus:
    PS C:\> .\CopyCreateNewUser.ps1 -TemplateUser "Vorlage01" -NewUserName "NeuerUser" -NewUserPassword "SicheresP@ssw0rd!123" -Verify

.EXAMPLE
    # Batch-Erstellung via CSV (Delimiter ;):
    PS C:\> .\CopyCreateNewUser.ps1 -CsvPath "C:\Pfad\zu\users.csv" -Verify

.NOTES
    Version: 4.0
    Autor: IT-Abteilung
    Letzte Änderung: 2024-10-15
#>

[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'Interactive')]
param (
    [Parameter(ParameterSetName = 'CSV')]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [string]$CsvPath,
    
    [Parameter(ParameterSetName = 'Single')]
    [string]$TemplateUser,
    
    [Parameter(ParameterSetName = 'Single')]
    [string]$NewUserName,
    
    [Parameter(ParameterSetName = 'Single')]
    [string]$NewUserPassword,
    
    [Parameter()]
    [switch]$Verify
)

#region Initialisierung & Hilfsfunktionen

# AD-Modul importieren (Abbruch bei Fehler)
try {
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch {
    Write-Error "Active Directory Modul konnte nicht geladen werden. Bitte RSAT installieren."
    exit 1
}

# Logging-Funktion – schreibt sowohl in die Konsole (mit Farbausgabe) als auch in eine Log-Datei
$LogFile = Join-Path $PSScriptRoot "ADUserCreation_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARNING","ERROR","SUCCESS")]
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$timestamp][$Level] $Message"
    Add-Content -Path $LogFile -Value $entry
    $color = switch ($Level) {
        "INFO"    { "Gray" }
        "WARNING" { "Yellow" }
        "ERROR"   { "Red" }
        "SUCCESS" { "Green" }
    }
    Write-Verbose $entry
    Write-Host $entry -ForegroundColor $color
}

# Testet, ob ein Kennwort komplex genug ist (Mindestlänge, Groß-/Kleinschreibung, Ziffern, Sonderzeichen)
function Test-PasswordComplexity {
    param([string]$Password)
    $rules = @(
        { $Password.Length -ge 12 },
        { $Password -match '[A-Z]' },
        { $Password -match '[a-z]' },
        { $Password -match '\d' },
        { $Password -match '[\W_]' }
    )
    foreach ($rule in $rules) {
        if (-not (& $rule)) { return $false }
    }
    return $true
}

# Kopiert Gruppenmitgliedschaften vom Quellbenutzer zum Zielbenutzer
function Copy-ADGroupMembership {
    param(
        [string]$SourceUser,
        [string]$TargetUser
    )
    try {
        # Hole die Gruppen (MemberOf) aus dem Template
        $groups = Get-ADUser -Identity $SourceUser -Properties MemberOf | Select-Object -ExpandProperty MemberOf
        foreach ($groupDN in $groups) {
            if ($PSCmdlet.ShouldProcess($groupDN, "Mitgliedschaft hinzufügen")) {
                Add-ADGroupMember -Identity $groupDN -Members $TargetUser -ErrorAction Stop
            }
        }
        Write-Log "Gruppenmitgliedschaften (Anzahl: $($groups.Count)) kopiert." -Level SUCCESS
    }
    catch {
        Write-Log "Fehler beim Kopieren der Gruppen: $_" -Level ERROR
        throw $_
    }
}

# Erstellt einen neuen AD-Benutzer anhand eines Vorlagenbenutzers,
# unterscheidet hierbei erforderliche (z.B. SAMAccountName, AccountPassword, Name)
# von optionalen Attributen (GivenName, Surname, Department etc.).
# Kopiert zusätzlich OU-Platzierung und weitere kopierbare AD-Eigenschaften.
function New-ADUserFromTemplate {
    param(
        [string]$TemplateSam,
        [string]$NewSam,
        [string]$Password,  # als Klartext, wird intern in SecureString konvertiert
        [hashtable]$OptionalAttributes
    )
    try {
        # Lade den Vorlagenbenutzer (alle kopierbaren Attribute)
        $template = Get-ADUser -Identity $TemplateSam -Properties *
        if (-not $template) {
            throw "Vorlagenbenutzer $TemplateSam nicht gefunden."
        }
        
        # Bestimme das Ziel-OU aus dem DistinguishedName der Vorlage
        $templateDNParts = $template.DistinguishedName -split ',',2
        $ouPath = $templateDNParts[1]  # z.B.: OU=Benutzer,DC=domain,DC=local
        
        # Konvertiere das Plaintext-Passwort in einen SecureString
        $securePass = ConvertTo-SecureString $Password -AsPlainText -Force
        
        # Erstelle ein Hashtable für die zwingend neu zu setzenden Parameter
        $userParams = @{
            SamAccountName        = $NewSam
            Name                  = $NewSam  # Standardmäßig: neuer Benutzername (kann evtl. in OptionalAttributes überschrieben werden)
            AccountPassword       = $securePass
            Enabled               = $true
            ChangePasswordAtLogon = $true
            Path                  = $ouPath
            UserPrincipalName     = "$NewSam@$((Get-ADDomain).DNSRoot)"
        }
        
        # Optional werden Attribute aus dem OptionalAttributes Hashtable übernommen.
        # Wichtige Attribute wie GivenName, Surname, DisplayName etc. können hier gesetzt werden.
        foreach ($key in $OptionalAttributes.Keys) {
            # Nur Attribute übernehmen, die nicht zwingend aus der Vorlage verwendet werden sollen, z.B. wenn sie individuell sein müssen:
            $userParams[$key] = $OptionalAttributes[$key]
        }
        
        # Das Klonen der Vorlageneigenschaften mit -Instance kopiert viele AD Attribute,
        # aber zwingend neu zu setzende Parameter (wie Kennwort, SAMAccountName etc.) werden überschrieben.
        if ($PSCmdlet.ShouldProcess($NewSam, "Erstelle neuen Benutzer basierend auf Vorlage $TemplateSam")) {
            $newUser = New-ADUser -Instance $template @userParams -PassThru -ErrorAction Stop
            Write-Log "Neuer Benutzer $NewSam wurde erstellt (OU: $ouPath)." -Level SUCCESS
            return $newUser
        }
    }
    catch {
        Write-Log "Fehler beim Erstellen des Benutzers $NewSam: $_" -Level ERROR
        return $null
    }
}
#endregion

#region Verarbeitungslogik

# Initialisiere leere Liste für erstellte Benutzer
$createdUsers = @()

if ($PSBoundParameters.ContainsKey("CsvPath")) {
    # ===============================
    # CSV-Batch-Modus (deutsche CSV: Delimiter ;)
    # ===============================
    Write-Log "Starte Batch-Erstellung via CSV: $CsvPath" -Level INFO
    try {
        $users = Import-Csv -Path $CsvPath -Delimiter ";" -Encoding UTF8
    }
    catch {
        Write-Error "CSV-Datei konnte nicht geladen werden: $_"
        exit 1
    }
    
    # Prüfe, ob alle Pflichtspalten vorhanden sind
    $requiredCols = @("TemplateUser", "NewUserName", "NewUserPassword")
    $csvCols = $users[0].PSObject.Properties.Name
    $missing = $requiredCols | Where-Object {$_ -notin $csvCols}
    if ($missing) {
        Write-Error "Fehlende Pflichtspalten in CSV: $($missing -join ', ')"
        exit 1
    }
    
    # Für jeden Eintrag in der CSV
    foreach ($u in $users) {
        Write-Log "Verarbeite Benutzer: $($u.NewUserName)" -Level INFO
        
        # Passwort-Komplexitätsprüfung
        if (-not (Test-PasswordComplexity -Password $u.NewUserPassword)) {
            Write-Log "Passwort für $($u.NewUserName) entspricht nicht den Richtlinien." -Level ERROR
            continue
        }
        
        # Optional: Erlaube Überschreibung einzelner Attribute (z. B. GivenName, Surname, Department, Office)
        $optAttrs = @{}
        if ($u.PSObject.Properties.Name -contains "GivenName") { $optAttrs["GivenName"] = $u.GivenName }
        if ($u.PSObject.Properties.Name -contains "Surname")  { $optAttrs["Surname"] = $u.Surname }
        if ($u.PSObject.Properties.Name -contains "DisplayName") { $optAttrs["DisplayName"] = $u.DisplayName }
        if ($u.PSObject.Properties.Name -contains "Department") { $optAttrs["Department"] = $u.Department }
        if ($u.PSObject.Properties.Name -contains "Office")     { $optAttrs["Office"] = $u.Office }
        
        # Ggf. kann in der CSV auch der TemplateUser überschrieben werden; ansonsten Standard:
        $templateForUser = if ($u.TemplateUser) { $u.TemplateUser } else { $TemplateUser }
        
        $newUser = New-ADUserFromTemplate -TemplateSam $templateForUser `
                                          -NewSam $u.NewUserName `
                                          -Password $u.NewUserPassword `
                                          -OptionalAttributes $optAttrs
        if ($newUser) {
            # Kopiere Gruppenmitgliedschaften vom Template
            try {
                Copy-ADGroupMembership -SourceUser $templateForUser -TargetUser $u.NewUserName
            }
            catch { Write-Log "Fehler beim Kopieren der Gruppen für $($u.NewUserName)" -Level ERROR }
            $createdUsers
