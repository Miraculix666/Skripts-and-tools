<#
.SYNOPSIS
    Erstellt neue AD-Benutzer als exakte Kopien eines Vorlagenbenutzers.

.DESCRIPTION
    Dieses Skript erstellt einen oder mehrere neue Active Directory Benutzer, indem es einen vorhandenen Vorlagenbenutzer kopiert.
    Es übernimmt alle kopierbaren Attribute, z.B. Gruppenmitgliedschaften, OU‑Platzierung, AD‑Rechte und optional weitere 
    benutzerdefinierte Werte. Dabei werden zwingend neue Attribute (SAMAccountName, AccountPassword, Name) von optionalen 
    Attributen (GivenName, Surname, DisplayName, Abteilung, Office, etc.) unterschieden.
    
    Das Skript unterstützt drei Aufrufvarianten:
      • Keine Parameter: Interaktiver Modus – es werden alle fehlenden Parameter abgefragt.
      • Einzelbenutzer: Alle nötigen Parameter (TemplateUser, NewUserName, NewUserPassword) werden übergeben.
      • CSV-Batch: Mit dem Parameter –CsvPath (Die CSV-Datei verwendet in der deutschen Version den Semikolon‑Delimiter).
        
    WICHTIG: Im CSV‑Modus müssen folgende Spalten vorhanden sein (mit Semikolon als Trenner):
        TemplateUser;NewUserName;NewUserPassword;GivenName;Surname;DisplayName;Department;Office;... (Weitere optionale Attribute)
        
    Das Skript wandelt Plaintext ‑ Kennwörter in SecureStrings um und gibt detaillierte Log‑ und “verbose” Ausgaben aus.

.PARAMETER CsvPath
    (Optional) Pfad zu einer CSV-Datei, in der die neuen Benutzer definiert sind. 
    Die CSV muss die Pflichtspalten TemplateUser, NewUserName und NewUserPassword enthalten,
    optional werden weitere Spalten wie GivenName, Surname, DisplayName, Department, Office etc. unterstützt.

.PARAMETER TemplateUser
    (Optional, außer CSV‑Modus) Der SAMAccountName des Vorlagenbenutzers.

.PARAMETER NewUserName
    (Optional, außer CSV‑Modus) Der SAMAccountName des neuen AD-Benutzers.

.PARAMETER NewUserPassword
    (Optional, außer CSV‑Modus) Das Plaintext‑Kennwort für den neuen Benutzer.
    Das Skript konvertiert es in einen SecureString.

.PARAMETER Verify
    Schalter zur Aktivierung einer abschließenden Verifikation der erstellten Benutzer (z.B. Attributvergleich und Gruppen).

.PARAMETER Verbose
    Schalter zur Aktivierung detaillierter Ausgaben während der Skriptausführung.

.EXAMPLE
    # Interaktiver Modus:
    PS C:\> .\CopyCreateNewUser.ps1 -Verbose

.EXAMPLE
    # Einzelbenutzermodus:
    PS C:\> .\CopyCreateNewUser.ps1 -TemplateUser "Vorlage01" -NewUserName "NeuerUser" -NewUserPassword "SicheresP@ssw0rd!123" -Verify -Verbose

.EXAMPLE
    # Batch-Erstellung via CSV (Delimiter ;):
    PS C:\> .\CopyCreateNewUser.ps1 -CsvPath "C:\Pfad\zu\users.csv" -Verify -Verbose

.NOTES
    Version: 5.1
    Autor: IT-Abteilung / basierend auf dem Script von Miraculix666 und den Beispielen von Petri.com
    Letzte Änderung: 2024-10-27
#>

[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'Interactive')]
param (
    [Parameter(ParameterSetName = 'CSV')]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [string]$CsvPath,
    
    [Parameter(ParameterSetName = 'Single')]
    [Parameter(ParameterSetName = 'Interactive')]
    [string]$TemplateUser,
    
    [Parameter(ParameterSetName = 'Single')]
    [Parameter(ParameterSetName = 'Interactive')]
    [string]$NewUserName,
    
    [Parameter(ParameterSetName = 'Single')]
    [Parameter(ParameterSetName = 'Interactive')]
    [string]$NewUserPassword,
    
    [Parameter()]
    [switch]$Verify,

    [Parameter()]
    [switch]$Verbose
)

#region Initialisierung & Hilfsfunktionen

# AD-Modul importieren (Abbruch bei Fehler)
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Verbose "Active Directory Modul erfolgreich geladen."
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
    Add-Content -Path $LogFile -Value $entry -Encoding UTF8
    $color = switch ($Level) {
        "INFO"    { "Gray" }
        "WARNING" { "Yellow" }
        "ERROR"   { "Red" }
        "SUCCESS" { "Green" }
    }
    if ($Verbose) {
        Write-Verbose $entry
    }
    Write-Host $entry -ForegroundColor $color
}

# Testet, ob ein Kennwort komplex genug ist (Mindestlänge, Groß-/Kleinschreibung, Ziffern, Sonderzeichen)
function Test-PasswordComplexity {
    param([string]$Password)
    Write-Verbose "Prüfe Passwortkomplexität für: $($Password -replace '.*','*')" # Maskiere das Passwort in der Ausgabe
    $rules = @(
        { $Password.Length -ge 12 }                # Mindestlänge 12 Zeichen
        { $Password -match '[A-Z]' }               # Mindestens ein Großbuchstabe
        { $Password -match '[a-z]' }               # Mindestens ein Kleinbuchstabe
        { $Password -match '\d' }                  # Mindestens eine Ziffer
        { $Password -match '[\W_]' }               # Mindestens ein Sonderzeichen
    )
    foreach ($rule in $rules) {
        if (-not (& $rule)) { 
            Write-Verbose "Passwortregel nicht erfüllt: $($rule -replace '^.*\{|\}.*$')"
            return $false 
        }
    }
    Write-Verbose "Passwortkomplexität ausreichend."
    return $true
}

# Kopiert Gruppenmitgliedschaften vom Quellbenutzer zum Zielbenutzer
function Copy-ADGroupMembership {
    param(
        [string]$SourceUser,
        [string]$TargetUser
    )
    try {
        Write-Verbose "Kopiere Gruppenmitgliedschaften von '$SourceUser' zu '$TargetUser'."
        # Hole die Gruppen (MemberOf) aus dem Template
        $groups = Get-ADUser -Identity $SourceUser -Properties MemberOf | Select-Object -ExpandProperty MemberOf
        Write-Verbose "Anzahl gefundener Gruppen: $($groups.Count)"
        foreach ($groupDN in $groups) {
            $groupName = (Get-ADGroup -Identity $groupDN).Name
            Write-Verbose "Füge '$TargetUser' zur Gruppe '$groupName' ($groupDN) hinzu."
            if ($PSCmdlet.ShouldProcess($groupName, "Mitgliedschaft hinzufügen")) {
                Add-ADGroupMember -Identity $groupDN -Members $TargetUser -ErrorAction Stop
                Write-Verbose "Benutzer '$TargetUser' erfolgreich zur Gruppe '$groupName' hinzugefügt."
            }
        }
        Write-Log "Gruppenmitgliedschaften (Anzahl: $($groups.Count)) von '$SourceUser' nach '$TargetUser' kopiert." -Level SUCCESS
    }
    catch {
        Write-Log "Fehler beim Kopieren der Gruppenmitgliedschaften von '$SourceUser' nach '$TargetUser': $_" -Level ERROR
        throw $_
    }
}

# Erstellt einen neuen AD-Benutzer anhand eines Vorlagenbenutzers.
# Kopiert OU-Platzierung und weitere kopierbare AD-Eigenschaften.
function New-ADUserFromTemplate {
    param(
        [string]$TemplateSam,
        [string]$NewSam,
        [string]$Password,  # als Klartext, wird intern in SecureString konvertiert
        [hashtable]$OptionalAttributes
    )
    try {
        Write-Verbose "Erstelle neuen Benutzer '$NewSam' basierend auf Vorlage '$TemplateSam'."
        # Lade den Vorlagenbenutzer (alle kopierbaren Attribute)
        $template = Get-ADUser -Identity $TemplateSam -Properties *
        if (-not $template) {
            throw "Vorlagenbenutzer '$TemplateSam' nicht gefunden."
        }
        Write-Verbose "Vorlagenbenutzer '$TemplateSam' gefunden."

        # Bestimme das Ziel-OU aus dem DistinguishedName der Vorlage
        $templateDNParts = $template.DistinguishedName -split ',', 2
        $ouPath = $templateDNParts[1]  # z.B.: OU=Benutzer,DC=domain,DC=local
        Write-Verbose "Ziel-OU für neuen Benutzer: '$ouPath'"

        # Konvertiere das Plaintext-Passwort in einen SecureString
        $securePass = ConvertTo-SecureString $Password -AsPlainText -Force
        Write-Verbose "Passwort erfolgreich in SecureString konvertiert."

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
        Write-Verbose "Zwingend zu setzende Attribute für neuen Benutzer festgelegt."

        # Optional werden Attribute aus dem OptionalAttributes Hashtable übernommen.
        Write-Verbose "Prüfe optionale Attribute."
        foreach ($key in $OptionalAttributes.Keys) {
            # Nur Attribute übernehmen, die nicht zwingend aus der Vorlage verwendet werden sollen (z.B. Name, SamAccountName, etc.)
            if (-not $userParams.ContainsKey($key)) {
                $userParams[$key] = $OptionalAttributes[$key]
                Write-Verbose "Übernehme optionales Attribut '$key' mit Wert '$($OptionalAttributes[$key])'."
            } else {
                Write-Verbose "Optionales Attribut '$key' wird nicht übernommen, da es zwingend gesetzt werden muss."
            }
        }

        # Das Klonen der Vorlageneigenschaften mit -Instance kopiert viele AD Attribute.
        Write-Verbose "Kopiere Attribute von Vorlagenbenutzer '$TemplateSam'."
        if ($PSCmdlet.ShouldProcess($NewSam, "Erstelle neuen Benutzer basierend auf Vorlage $TemplateSam")) {
            $newUser = New-ADUser -Instance $template @userParams -PassThru -ErrorAction Stop
            Write-Log "Neuer Benutzer '$NewSam' wurde erfolgreich erstellt (OU: '$ouPath')." -Level SUCCESS
            return $newUser
        }
    }
    catch {
        Write-Log "Fehler beim Erstellen des Benutzers `$NewSam`: $($Error[0].Exception.Message)" -Level ERROR
        return $null
    }
}

# Funktion zur Verifikation der Benutzererstellung
function Verify-UserCreation {
    param(
        [string]$UserName,
        [string]$TemplateUser
    )
    
    Write-Log "Starte Verifikation für Benutzer: $UserName" -Level INFO

    try {
        $newUser = Get-ADUser -Identity $UserName -Properties *
        $templateUser = Get-ADUser -Identity $TemplateUser -Properties *

        if ($newUser -and $templateUser) {
            Write-Log "Verifikation: Benutzer '$UserName' gefunden." -Level INFO
            Write-Log "Verifikation: Vorlagenbenutzer '$TemplateUser' gefunden." -Level INFO

            # Attributvergleich
            Write-Log "Starte Attributvergleich..." -Level INFO
            $userProperties = $newUser.PSObject.Properties | Where-Object {$_.MemberType -eq "Property"}
            $templateProperties = $templateUser.PSObject.Properties | Where-Object {$_.MemberType -eq "Property"}

            foreach ($property in $templateProperties) {
                if ($property.Name -notin ("SamAccountName", "Name", "ObjectGUID", "SID", "DistinguishedName", "UserPrincipalName", "PasswordLastSet", "PasswordNeverExpires", "Enabled", "msDS-UserPasswordExpiryTimeComputed")) {
                    $newValue = $userProperties | Where-Object {$_.Name -eq $property.Name} | Select-Object -ExpandProperty Value
                    $templateValue = $property.Value

                    if ($newValue -ne <span class="math-inline">templateValue\) \{
Write\-Log "Attribut '</span>($property.Name)' stimmt nicht überein. Vorlage: '$templateValue', Benutzer: '<span class="math-inline">newValue'" \-Level WARNING
\} else \{
Write\-Log "Attribut '</span>($property.Name)' stimmt überein." -Level SUCCESS
                    }
                }
            }
            
            # Gruppenmitgliedschaften überprüfen
            Write-Log "Überprüfe Gruppenmitgliedschaften..." -Level INFO
            $userGroups = Get-ADUser $UserName -Properties MemberOf | Select-Object -ExpandProperty MemberOf
            $templateGroups = Get-ADUser $TemplateUser -Properties MemberOf | Select-Object -ExpandProperty MemberOf
            
            $missingGroups = Compare-Object -ReferenceObject $templateGroups -DifferenceObject $userGroups | Where-Object {$_.SideIndicator -eq "<="}
            
            if ($missingGroups) {
                foreach ($group in $missingGroups) {
                    $groupName = (Get-ADGroup -Identity $group.InputObject).Name
                    Write-Log "Benutzer '$UserName' fehlt in Gruppe '$groupName' (Vorlage)." -Level WARNING
                }
            } else {
                Write-Log "Gruppenmitgliedschaften stimmen überein." -Level SUCCESS
            }

        } else {
            Write-Log "Verifikation: Benutzer '$UserName' oder Vorlagenbenutzer '$TemplateUser' nicht gefunden." -Level ERROR
        }
    }
    catch {
        Write-Log "Fehler während der Verifikation: $_" -Level ERROR
    }
}

#endregion

#region Verarbeitungslogik

# Initialisiere leere Liste für erstellte Benutzer
$createdUsers = @()

# Bestimmung des aktuellen Parametersatzes und Abfrage fehlender Parameter im interaktiven Modus
$paramSetName = $PSCmdlet.ParameterSetName
Write-Verbose "Aktueller Parametersatz: $paramSetName"

if ($paramSetName -eq 'Interactive') {
    Write-Log "Starte interaktiven Modus." -Level INFO
    
    if (-not $TemplateUser) {
        $TemplateUser = Read-Host "Bitte geben Sie den SAMAccountName des Vorlagenbenutzers ein"
    }
    if (-not $NewUserName) {
        $NewUserName = Read-Host "Bitte geben Sie den SAMAccountName des neuen Benutzers ein"
    }
    if (-not $NewUserPassword) {
        $NewUserPassword = Read-Host "Bitte geben Sie das Passwort für den neuen Benutzer ein"
    }
}

# Wenn der Parameter 'TemplateUser' übergeben wurde, aber nicht 'NewUserName' oder 'NewUserPassword',
# dann wechsle in den interaktiven Modus für die fehlenden Parameter.
if ($TemplateUser -and (-not $NewUserName -or -not $NewUserPassword)) {
    Write-Log "Wechsle in den interaktiven Modus für fehlende Parameter." -Level INFO
    
    if (-not $NewUserName) {
        $NewUserName = Read-Host "Bitte geben Sie den SAMAccountName des neuen Benutzers ein"
    }
    if (-not $NewUserPassword) {
        $NewUserPassword = Read-Host "Bitte geben Sie das Passwort für den neuen Benutzer ein"
    }
}

# Validierung der Eingabeparameter, falls vorhanden
if ($TemplateUser) {
    if (-not (Get-ADUser -F {SamAccountName -eq $TemplateUser})) {
        Write-Error "Vorlagenbenutzer '$TemplateUser' existiert nicht."
        exit 1
    }
    Write-Verbose "Vorlagenbenutzer '$TemplateUser' validiert."
}

if ($NewUserName) {
    if (Get-ADUser -F {SamAccountName -eq $NewUserName}) {
        Write-Error "Benutzer '$NewUserName' existiert bereits."
        exit 1
    }
    Write-Verbose "Neuer Benutzername '$NewUserName' ist verfügbar."
}

if ($NewUserPassword) {
    if (-not (Test-PasswordComplexity -Password $NewUserPassword)) {
        Write-Error "Passwort entspricht nicht den Komplexitätsanforderungen."
        exit 1
    }
    Write-Verbose "Passwort validiert."
}


if ($PSBoundParameters.ContainsKey("CsvPath")) {
    # ===============================
    # CSV-Batch-Modus (deutsche CSV: Delimiter ;)
    # ===============================
    Write-Log "Starte Batch-Erstellung via CSV: '$CsvPath'" -Level INFO
    try {
        # Importiere CSV mit deutschem Lokalisierungsprofil (Semikolon-Trenner, UTF8-Encoding)
        $users = Import
