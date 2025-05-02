#Enhanced-AD-User-Managment.ps1
<#
.SYNOPSIS
Verwaltet Active Directory-Benutzer: Kopiert einzelne Benutzer oder erstellt Benutzer aus einer CSV-Datei.

.DESCRIPTION
Dieses Skript bietet Werkzeuge zur Verwaltung von Active Directory (AD) Benutzern.
Es kann verwendet werden, um:
- Einen einzelnen AD-Benutzer zu kopieren, einschließlich Gruppenzugehörigkeiten und optional der OU-Struktur (Modus: CopySingleUser).
- Mehrere AD-Benutzer basierend auf einer CSV-Datei zu erstellen (Modus: CreateUsersFromCSV). Optional können Attribute und Gruppen von einem Referenzbenutzer (Vorlage) übernommen werden.

Das Skript ist für PowerShell Version 5.1 optimiert und verwendet deutsche Lokalisierungseinstellungen für CSV-Exporte und Datums-/Zeitformate.
Es implementiert detaillierte Protokollierung und unterstützt den -Verbose Parameter für ausführliche Ausgaben.
Standardmäßig werden Log-Dateien und Benutzerberichte im Verzeichnis des Skripts gespeichert.

.PARAMETER CopySingleUser
Schalter, um den Modus zum Kopieren eines einzelnen Benutzers zu aktivieren.

.PARAMETER ReferenceUserSamAccountName
Der SAMAccountName des Referenzbenutzers.
- Im Modus 'CopySingleUser': Der Quellbenutzer, der kopiert werden soll (Mandatory).
- Im Modus 'CreateUsersFromCSV': Der Vorlagenbenutzer, von dem optional Attribute/Gruppen übernommen werden (Optional).

.PARAMETER TargetUserSamAccountName
Der gewünschte SAMAccountName für den neuen (kopierten) Benutzer (nur für Modus CopySingleUser). Wird interaktiv abgefragt, wenn nicht angegeben.

.PARAMETER TargetUserPassword
Das initiale Passwort für den neuen (kopierten) Benutzer als SecureString (nur für Modus CopySingleUser). Wird interaktiv und sicher abgefragt, wenn nicht angegeben. Der Benutzer muss das Passwort bei der ersten Anmeldung ändern.

.PARAMETER TargetOU
Der Distinguished Name der Organisationseinheit (OU), in die der neue Benutzer kopiert/erstellt werden soll.
Für CopySingleUser: Optional; Standard ist die OU des Referenzbenutzers.
Für CreateUsersFromCSV: Optional; Überschreibt die OU aus der CSV oder vom Referenzbenutzer.

.PARAMETER Force
Überschreibt einen bereits existierenden Zielbenutzer im Modus CopySingleUser, falls vorhanden. Standardmäßig wird ein Fehler ausgegeben.

.PARAMETER CreateUsersFromCSV
Schalter, um den Modus zur Erstellung von Benutzern aus einer CSV-Datei zu aktivieren.

.PARAMETER CsvPath
Der vollständige Pfad zur CSV-Datei, die die Daten für die zu erstellenden Benutzer enthält (nur für Modus CreateUsersFromCSV).
Erwartete Spalten (mindestens): SamAccountName, GivenName, Surname.
Optionale Spalten: Password (Klartext - NICHT EMPFOHLEN!), TargetOU, Enabled, Description, Office, Department, Title, Company, EmailAddress, StreetAddress, City, State, PostalCode, Country, OfficePhone. Siehe .NOTES für Details.

.PARAMETER DefaultPassword
Ein Standardpasswort (als SecureString), das für alle Benutzer aus der CSV verwendet wird, es sei denn, die CSV enthält eine 'Password'-Spalte. Wenn weder DefaultPassword noch eine 'Password'-Spalte vorhanden sind, wird ein zufälliges Passwort generiert (empfohlen).

.PARAMETER LogPath
Verzeichnis für die Log-Dateien. Standard ist das Verzeichnis, in dem das Skript liegt (`$PSScriptRoot`), oder das aktuelle Arbeitsverzeichnis (`$PWD`), wenn nicht aus einer Datei ausgeführt.

.PARAMETER LogLevel
Steuert die Detailtiefe der Log-Datei. Mögliche Werte: Error, Warning, Info, Verbose. Standard ist 'Info'.

.PARAMETER UserReportCsvPath
Pfad für die CSV-Datei, die einen Bericht über erstellte/kopierte Benutzer enthält. Standard ist das Skriptverzeichnis (`$PSScriptRoot` oder `$PWD`) mit dem Namen '{ScriptName}_UserReport_{Timestamp}.csv'.

.PARAMETER NoUserReport
Unterdrückt die Erstellung des Benutzerberichts.

.EXAMPLE
# Beispiel 1: Kopiert 'BenutzerA' zu 'BenutzerB' interaktiv (fragt nach Zielname und Passwort)
.\Enhanced-ADManagement.ps1 -CopySingleUser -ReferenceUserSamAccountName BenutzerA

.EXAMPLE
# Beispiel 2: Kopiert 'BenutzerA' zu 'BenutzerC', setzt Passwort, legt in spezifischer OU ab und überschreibt Ziel falls vorhanden
$password = ConvertTo-SecureString "P@sswOrd123!" -AsPlainText -Force
.\Enhanced-ADManagement.ps1 -CopySingleUser -ReferenceUserSamAccountName BenutzerA -TargetUserSamAccountName BenutzerC -TargetUserPassword $password -TargetOU "OU=NeueMitarbeiter,DC=firma,DC=local" -Force -Verbose

.EXAMPLE
# Beispiel 3: Erstellt Benutzer aus CSV mit Standardpasswort und Gruppen/Attributen von 'TemplateUser'
$defaultPass = ConvertTo-SecureString "Sommer2025!" -AsPlainText -Force
.\Enhanced-ADManagement.ps1 -CreateUsersFromCSV -CsvPath "C:\temp\neue_benutzer.csv" -ReferenceUserSamAccountName TemplateUser -DefaultPassword $defaultPass -LogLevel Verbose

.EXAMPLE
# Beispiel 4: Erstellt Benutzer aus CSV, verwendet Passwort aus CSV (WARNUNG: Unsicher!) und speichert Log in C:\Logs
.\Enhanced-ADManagement.ps1 -CreateUsersFromCSV -CsvPath "C:\temp\neue_benutzer_mit_passwort.csv" -TargetOU "OU=Vertrieb,DC=firma,DC=local" -LogPath "C:\Logs" -NoUserReport

.NOTES
Autor: Gemini (basierend auf Nutzer-Input und Beispielen)
Version: 2.0
Datum: 2025-05-02
Benötigte Module: ActiveDirectory (wird durch #requires geprüft)
Benötigte Berechtigungen: Ausreichende AD-Berechtigungen zum Lesen von Benutzern und zum Erstellen/Modifizieren von Benutzern. Schreibrechte im Zielverzeichnis für Logs/Berichte.

CSV-Format für CreateUsersFromCSV:
- Trennzeichen: Semikolon (;)
- Kodierung: UTF8
- Erforderliche Spalten: SamAccountName, GivenName (Vorname), Surname (Nachname)
- Empfohlene Spalten für volle Funktionalität: EmailAddress, TargetOU (Distinguished Name)
- Optionale Spalten (werden verwendet, wenn vorhanden): Password (Klartext - NICHT EMPFOHLEN!), Enabled (true/false), Description, Office, Department, Title, Company, StreetAddress, City, State, PostalCode, Country, OfficePhone.
- Wenn die Spalte 'Password' nicht vorhanden ist oder leer ist, wird -DefaultPassword verwendet. Wenn beides fehlt, wird ein sicheres, zufälliges Passwort generiert und MUSS geändert werden.
- Wenn die Spalte 'TargetOU' nicht vorhanden ist oder leer ist, wird der Wert von -TargetOU verwendet. Wenn dieser auch fehlt, wird die OU des ReferenceUser (falls angegeben) verwendet. Ansonsten Fehler.
- Wenn die Spalte 'Enabled' nicht vorhanden ist, wird der Benutzer standardmäßig aktiviert ($true).

Passwort-Sicherheit: Das Speichern von Klartext-Passwörtern in CSV-Dateien ist ein erhebliches Sicherheitsrisiko! Verwenden Sie bevorzugt den Parameter -DefaultPassword oder generieren Sie zufällige Passwörter und kommunizieren Sie diese sicher.

Standardpfade: Wenn -LogPath oder -UserReportCsvPath nicht angegeben werden, versucht das Skript, die Dateien im selben Verzeichnis wie das Skript selbst (`$PSScriptRoot`) zu speichern. Wenn das Skript nicht aus einer Datei ausgeführt wird (z.B. im ISE oder direkt in der Konsole), wird stattdessen das aktuelle Arbeitsverzeichnis (`$PWD`) verwendet. Stellen Sie sicher, dass Schreibberechtigungen im Zielverzeichnis vorhanden sind.

Testen: Führen Sie das Skript zuerst in einer Testumgebung aus!

Quellen/Referenzen (vom Benutzer bereitgestellt):
https://petri.com/how-to-copy-active-directory-groups-from-one-user-to-another-with-powershell/
https://petri.com/how-to-copy-active-directory-users-with-powershell/
https://petri.com/create-new-active-directory-users-excel-powershell
https://blog.netwrix.com/bulk-user-creation-in-active-directory/
https://learn.microsoft.com/en-us/answers/questions/1035531/creating-new-ad-users-in-powershell-from-existing?page=2#answers
https://github.com/RichPrescott/UserCreation/blob/master/ANUC.ps1
https://community.spiceworks.com/t/bulk-create-active-directory-users-powershell-with-logs-less-rows-in-csv/974593
https://forums.powershell.org/t/copy-ad-user-not-from-template-just-mirror/18842/8
https://support.microsoft.com/en-us/windows/manage-user-accounts-in-windows-104dc19f-6430-4b49-6a2b-e4dbd1dcdf32
https://activedirectorypro.com/copy-group-membership-from-one-user-to-another-in-ad/
https://community.spiceworks.com/t/powershell-add-users-error/609976
https://community.spiceworks.com/t/how-to-copy-group-membership-from-one-user-to-another/1014279
https://notebooklm.google.com/notebook/9f6821c0-f64d-4adf-a462-c68fd050aea4?_gl=1*1p5hmk1*_ga*MTkyNjMzNzExNC4xNzM4NzYzMjU2*_ga_W0LDH41ZCB*MTc0NjE4MDcwMi4xMC4xLjE3NDYxODA3MDIuMC4wLjA.

.COMPONENT
ActiveDirectory

.ROLE
Administrator

.FUNCTIONALITY
User Account Management
#>

#requires -Version 5.1
#requires -Modules ActiveDirectory

[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'CopySingleUser')]
param(
    # --- Modus Schalter ---
    [Parameter(ParameterSetName = 'CopySingleUser', Mandatory = $true, HelpMessage = "Aktiviert den Modus zum Kopieren eines einzelnen Benutzers.")]
    [switch]$CopySingleUser,

    [Parameter(ParameterSetName = 'CreateUsersFromCSV', Mandatory = $true, HelpMessage = "Aktiviert den Modus zur Benutzererstellung aus CSV.")]
    [switch]$CreateUsersFromCSV,

    # --- Gemeinsame Parameter ---
    [Parameter(ParameterSetName = 'CopySingleUser', Mandatory = $true, HelpMessage = "SAMAccountName des Quellbenutzers für die Kopie.")]
    [Parameter(ParameterSetName = 'CreateUsersFromCSV', Mandatory = $false, HelpMessage = "SAMAccountName des Vorlagenbenutzers für Attribute/Gruppen.")]
    [ValidateNotNullOrEmpty()]
    [string]$ReferenceUserSamAccountName,

    [Parameter(ParameterSetName = 'CopySingleUser')]
    [Parameter(ParameterSetName = 'CreateUsersFromCSV')]
    [ValidateNotNullOrEmpty()]
    [string]$TargetOU, # Gilt für beide ParameterSets

    # --- Parameter für CopySingleUser ---
    [Parameter(ParameterSetName = 'CopySingleUser', Mandatory = $false, HelpMessage = "Gewünschter SAMAccountName für den neuen Benutzer.")]
    [ValidateNotNullOrEmpty()]
    [string]$TargetUserSamAccountName,

    [Parameter(ParameterSetName = 'CopySingleUser', Mandatory = $false, HelpMessage = "Initiales Passwort für den neuen Benutzer (SecureString).")]
    [System.Security.SecureString]$TargetUserPassword,

    [Parameter(ParameterSetName = 'CopySingleUser', HelpMessage = "Überschreibt einen existierenden Zielbenutzer.")]
    [switch]$Force,

    # --- Parameter für CreateUsersFromCSV ---
    [Parameter(ParameterSetName = 'CreateUsersFromCSV', Mandatory = $true, HelpMessage = "Pfad zur CSV-Datei.")]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$CsvPath,

    [Parameter(ParameterSetName = 'CreateUsersFromCSV', Mandatory = $false, HelpMessage = "Standardpasswort (SecureString) für CSV-Benutzer.")]
    [System.Security.SecureString]$DefaultPassword,

    # --- Globale Parameter ---
    [Parameter(Mandatory = $false, HelpMessage = "Verzeichnis für Log-Dateien. Standard: Skriptverzeichnis oder aktuelles Verzeichnis.")]
    [string]$LogPath,

    [Parameter(Mandatory = $false, HelpMessage = "Detailtiefe der Log-Datei (Error, Warning, Info, Verbose).")]
    [ValidateSet('Error', 'Warning', 'Info', 'Verbose')]
    [string]$LogLevel = 'Info',

    [Parameter(Mandatory = $false, HelpMessage = "Pfad für den CSV-Benutzerbericht. Standard: Skriptverzeichnis oder aktuelles Verzeichnis.")]
    [string]$UserReportCsvPath,

    [Parameter(Mandatory = $false, HelpMessage = "Unterdrückt die Erstellung des Benutzerberichts.")]
    [switch]$NoUserReport
)

begin {
    # --- Initialisierungen ---
    Write-Verbose "Beginne Initialisierung des Skripts."

    # Setze Kultur auf Deutsch für korrekte Formatierungen (z.B. CSV-Trennzeichen)
    try {
        [System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo('de-DE')
        [System.Threading.Thread]::CurrentThread.CurrentUICulture = [System.Globalization.CultureInfo]::GetCultureInfo('de-DE')
        Write-Verbose "Kultur auf 'de-DE' gesetzt."
    }
    catch {
        Write-Warning "Konnte Kultur nicht auf 'de-DE' setzen: $_. Standardeinstellungen werden verwendet."
    }

    # Fehlerbehandlung standardmäßig auf Stop setzen
    $ErrorActionPreference = 'Stop'

    # Bestimme Basisverzeichnis für Logs/Reports
    $basePath = $PSScriptRoot # Bevorzugt Skriptverzeichnis
    if (-not $basePath) {
        $basePath = $PWD.Path # Fallback: Aktuelles Arbeitsverzeichnis
        Write-Verbose "Variable `$PSScriptRoot ist leer. Verwende aktuelles Arbeitsverzeichnis als Basis für Standardpfade: $basePath"
    } else {
        Write-Verbose "Verwende Skriptverzeichnis als Basis für Standardpfade: $basePath"
    }

    # Bestimme Skriptnamen für Dateipräfixe
    $scriptBaseName = 'Enhanced-ADManagement' # Standard, falls nicht ermittelbar
    try {
        if ($MyInvocation.MyCommand.Name) {
            $scriptBaseName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
        }
    } catch {
        Write-Warning "Konnte Skriptnamen nicht automatisch ermitteln. Verwende '$scriptBaseName' als Präfix."
    }
    Write-Verbose "Verwende '$scriptBaseName' als Präfix für Log-/Berichtsdateien."


    # Logging Setup
    $scriptStartTime = Get-Date
    # Standard-Logpfad setzen, wenn nicht angegeben
    if (-not $PSBoundParameters.ContainsKey('LogPath')) {
        $LogPath = $basePath
    }
    $logFileName = "{0}_{1}_{2}.log" -f $scriptBaseName, $PSCmdlet.ParameterSetName, $scriptStartTime.ToString('yyyyMMdd-HHmmss')
    try {
        # Sicherstellen, dass das Zielverzeichnis existiert
        if (-not (Test-Path $LogPath -PathType Container)) {
            Write-Verbose "Erstelle Log-Verzeichnis: $LogPath"
            New-Item -Path $LogPath -ItemType Directory -Force:$true -ErrorAction Stop | Out-Null
        }
        $global:fullLogPath = Join-Path -Path $LogPath -ChildPath $logFileName # Global machen für Zugriff in Funktionen/End-Block
        Write-Verbose "Log-Datei wird sein: $fullLogPath"
        # Schreibe initialen Log-Eintrag
         "[$($scriptStartTime.ToString('yyyy-MM-dd HH:mm:ss'))] [Info] Skript '$scriptBaseName.ps1' gestartet (Modus: $($PSCmdlet.ParameterSetName)). LogLevel: $LogLevel. Ausgeführt von: $($env:USERNAME)." | Out-File -FilePath $fullLogPath -Encoding UTF8 -Append
    }
    catch {
        Write-Error "Fehler beim Initialisieren des Loggings nach '$LogPath': $_. Breche Skript ab."
        # Da Logging essentiell ist, hier wirklich abbrechen
        exit 1
    }

    # User Report Setup
    $global:userReportData = [System.Collections.Generic.List[PSObject]]::new() # Global für Sammlung über Modi hinweg
    $global:fullUserReportPath = $null
    if (-not $NoUserReport) {
        # Standard-Reportpfad setzen, wenn nicht angegeben
        if (-not $PSBoundParameters.ContainsKey('UserReportCsvPath')) {
            $UserReportCsvPath = $basePath
        }
         $reportFileName = "{0}_UserReport_{1}.csv" -f $scriptBaseName, $scriptStartTime.ToString('yyyyMMdd-HHmmss')
         try {
             # Sicherstellen, dass das Zielverzeichnis existiert
             if (-not (Test-Path $UserReportCsvPath -PathType Container)) {
                 Write-Verbose "Erstelle Berichts-Verzeichnis: $UserReportCsvPath"
                 New-Item -Path $UserReportCsvPath -ItemType Directory -Force:$true -ErrorAction Stop | Out-Null
             }
             $global:fullUserReportPath = Join-Path -Path $UserReportCsvPath -ChildPath $reportFileName
             Write-Verbose "Benutzerbericht wird erstellt: $fullUserReportPath"
         } catch {
             Write-Warning "Fehler beim Initialisieren des Benutzerberichts-Pfades '$UserReportCsvPath': $_. Bericht wird nicht erstellt."
             $NoUserReport = $true # Deaktiviere Berichterstellung bei Fehler
         }
    } else {
        Write-Verbose "Benutzerbericht wird nicht erstellt (-NoUserReport angegeben)."
    }


    # --- Hilfsfunktionen ---

    # Funktion zum Schreiben von Log-Einträgen
    function Write-Log {
        param(
            [Parameter(Mandatory = $true)]
            [ValidateSet('Error', 'Warning', 'Info', 'Verbose')]
            [string]$Level,

            [Parameter(Mandatory = $true)]
            [string]$Message
        )

        # Bestimme, ob die Nachricht basierend auf $LogLevel geloggt werden soll
        $logLevels = @{'Error' = 1; 'Warning' = 2; 'Info' = 3; 'Verbose' = 4 }
        $currentLogLevelValue = $logLevels[$LogLevel]
        $messageLogLevelValue = $logLevels[$Level]

        if ($messageLogLevelValue -le $currentLogLevelValue) {
            $logEntry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
            try {
                # Verwende den globalen Pfad
                Add-Content -Path $global:fullLogPath -Value $logEntry -Encoding UTF8 -ErrorAction Stop
            }
            catch {
                # Kritischer Fehler, wenn Logging fehlschlägt
                Write-Error "KRITISCH: Konnte Log-Eintrag nicht in '$($global:fullLogPath)' schreiben: $Message - Fehler: $_"
            }
        }

        # Zusätzliche Ausgabe auf der Konsole je nach Level
        switch ($Level) {
            'Error'   { Write-Error $Message }
            'Warning' { Write-Warning $Message }
            'Info'    { Write-Host "[INFO] $Message" -ForegroundColor Green } # Info messages green for visibility
            'Verbose' { Write-Verbose $Message } # Write-Verbose handles its own output based on $VerbosePreference/-Verbose switch
        }
    }

    # Funktion zum Hinzufügen von Daten zum Benutzerbericht
    function Add-UserReportEntry {
        param(
            [Parameter(Mandatory = $true)]
            [string]$SamAccountName,
            [Parameter(Mandatory = $true)]
            [string]$Status, # z.B. "Erstellt", "Kopiert", "Fehler"
            [Parameter(Mandatory = $false)]
            [string]$Detail = ""
        )
        if (-not $NoUserReport) {
            $reportObject = [PSCustomObject]@{
                Timestamp      = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                SamAccountName = $SamAccountName
                Status         = $Status
                Detail         = $Detail
            }
            $global:userReportData.Add($reportObject)
        }
    }


    # Funktion zum Kopieren eines AD-Benutzers mit Gruppen und OU
    function Copy-ADUserAdvanced {
        [CmdletBinding(SupportsShouldProcess = $true)]
        param(
            [Parameter(Mandatory = $true)]
            [Microsoft.ActiveDirectory.Management.ADUser]$SourceUser, # Ist der ReferenceUser

            [Parameter(Mandatory = $true)]
            [string]$TargetSamAccountName,

            [Parameter(Mandatory = $true)]
            [System.Security.SecureString]$Password,

            [Parameter(Mandatory = $false)]
            [string]$DestinationOU, # Optional: Ziel-OU

            [Parameter(Mandatory = $false)]
            [switch]$OverwriteTarget # Optional: Ziel überschreiben
        )

        Write-Log -Level Info -Message "Beginne Kopiervorgang von $($SourceUser.SamAccountName) nach $TargetSamAccountName."
        $targetUserExists = $false
        $existingTargetUser = $null
        try {
            $existingTargetUser = Get-ADUser -Filter "SamAccountName -eq '$TargetSamAccountName'" -ErrorAction SilentlyContinue
            if ($existingTargetUser) {
                $targetUserExists = $true
            }
        } catch {
             # Fehler beim Suchen ignorieren, weitermachen
            Write-Log -Level Warning -Message "Fehler beim Prüfen, ob Zielbenutzer '$TargetSamAccountName' existiert: $_"
        }

        if ($targetUserExists) {
            if (-not $OverwriteTarget) {
                $msg = "Zielbenutzer '$TargetSamAccountName' existiert bereits. Verwenden Sie -Force zum Überschreiben."
                Write-Log -Level Error -Message $msg
                Add-UserReportEntry -SamAccountName $TargetSamAccountName -Status "Fehler" -Detail $msg
                return $null # Fehler signalisieren
            } else {
                Write-Log -Level Warning -Message "Zielbenutzer '$TargetSamAccountName' existiert und wird überschrieben (-Force)."
                if ($PSCmdlet.ShouldProcess($TargetSamAccountName, "Vorhandenen Benutzer entfernen")) {
                    try {
                        Remove-ADUser -Identity $existingTargetUser -Confirm:$false -ErrorAction Stop
                        Write-Log -Level Info -Message "Vorhandener Benutzer '$TargetSamAccountName' entfernt."
                    } catch {
                        $msg = "Fehler beim Entfernen des vorhandenen Benutzers '$TargetSamAccountName': $_"
                        Write-Log -Level Error -Message $msg
                        Add-UserReportEntry -SamAccountName $TargetSamAccountName -Status "Fehler" -Detail $msg
                        return $null
                    }
                } else {
                    $msg = "Entfernen des vorhandenen Benutzers '$TargetSamAccountName' übersprungen (ShouldProcess)."
                    Write-Log -Level Info -Message $msg
                    Add-UserReportEntry -SamAccountName $TargetSamAccountName -Status "Übersprungen" -Detail $msg
                    return $null
                }
            }
        }

        # OU bestimmen
        $finalOU = $DestinationOU # Aus Parameter verwenden
        if (-not $finalOU) {
            # Wenn nicht im Parameter, nimm die OU des Quellbenutzers
            $finalOU = ($SourceUser.DistinguishedName -split ',', 2)[1]
            Write-Verbose "Keine Ziel-OU (-TargetOU) angegeben, verwende Quell-OU: $finalOU"
        }

        # Prüfen, ob die Ziel-OU existiert
        try {
            if (-not (Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$finalOU'" -ErrorAction Stop)) {
                 # Dies sollte eigentlich nicht passieren, da Get-ADOrganizationalUnit einen Fehler werfen sollte, wenn nicht gefunden. Doppelte Sicherheit.
                $msg = "Die Ziel-OU '$finalOU' existiert nicht."
                Write-Log -Level Error -Message $msg
                Add-UserReportEntry -SamAccountName $TargetSamAccountName -Status "Fehler" -Detail $msg
                return $null
            }
             Write-Verbose "Ziel-OU '$finalOU' ist gültig."
        } catch {
            $msg = "Fehler beim Überprüfen der Ziel-OU '$finalOU': $_"
            Write-Log -Level Error -Message $msg
            Add-UserReportEntry -SamAccountName $TargetSamAccountName -Status "Fehler" -Detail $msg
            return $null
        }


        # Parameter für New-ADUser vorbereiten
        # Wichtige Attribute, die NICHT direkt kopiert werden sollten oder müssen:
        # objectGUID, objectSid, sIDHistory, distinguishedName, lastLogon, lastLogonTimestamp, pwdLastSet, logonCount etc.
        # New-ADUser mit -Instance kopiert viele, aber wir setzen die wichtigsten explizit.
        $newUserParams = @{
            SamAccountName        = $TargetSamAccountName
            Name                  = $TargetSamAccountName # Standard Name = SamAccountName, kann später angepasst werden
            GivenName             = $SourceUser.GivenName
            Surname               = $SourceUser.Surname
            DisplayName           = "$($SourceUser.GivenName) $($SourceUser.Surname)" # Oder eine andere Logik
            UserPrincipalName     = "$TargetSamAccountName@$($env:USERDNSDOMAIN)" # Domain anpassen falls nötig!
            Path                  = $finalOU
            AccountPassword       = $Password
            ChangePasswordAtLogon = $true
            Enabled               = $true # Kopierte Benutzer standardmäßig aktivieren
            Description           = $SourceUser.Description # Beispiel für weitere Attribute
            Office                = $SourceUser.Office
            Department            = $SourceUser.Department
            Company               = $SourceUser.Company
            Title                 = $SourceUser.Title
            # Fügen Sie hier weitere Attribute hinzu, die kopiert werden sollen
            # Beispiel: -OfficePhone = $SourceUser.OfficePhone
        }

        # Benutzer erstellen
        $newUser = $null
        if ($PSCmdlet.ShouldProcess($TargetSamAccountName, "Neuen AD-Benutzer erstellen (Kopie von $($SourceUser.SamAccountName))")) {
            try {
                Write-Log -Level Info -Message "Erstelle Benutzer '$TargetSamAccountName' in OU '$finalOU'."
                # Verwende NICHT -Instance hier, um mehr Kontrolle zu haben
                $newUser = New-ADUser @newUserParams -PassThru -ErrorAction Stop
                Write-Log -Level Info -Message "Benutzer '$($newUser.SamAccountName)' erfolgreich erstellt (SID: $($newUser.SID.Value))."
                Add-UserReportEntry -SamAccountName $newUser.SamAccountName -Status "Kopiert" -Detail "Von $($SourceUser.SamAccountName) nach OU '$finalOU'"
            }
            catch {
                $msg = "Fehler beim Erstellen des Benutzers '$TargetSamAccountName': $_"
                Write-Log -Level Error -Message $msg
                Add-UserReportEntry -SamAccountName $TargetSamAccountName -Status "Fehler" -Detail $msg
                return $null
            }
        } else {
            $msg = "Erstellung von '$TargetSamAccountName' übersprungen (ShouldProcess)."
            Write-Log -Level Info -Message $msg
            Add-UserReportEntry -SamAccountName $TargetSamAccountName -Status "Übersprungen" -Detail $msg
            return $null
        }

        # Gruppenmitgliedschaften kopieren
        try {
            $sourceGroups = Get-ADPrincipalGroupMembership -Identity $SourceUser -ErrorAction Stop
             # Filter optionale problematische Gruppen (z.B. 'Domain Users' wird oft automatisch hinzugefügt)
             $groupsToCopy = $sourceGroups | Where-Object {$_.Name -ne "Domain Users"} # Beispiel Filter

            if ($groupsToCopy) {
                Write-Log -Level Info -Message "Kopiere $($groupsToCopy.Count) Gruppenmitgliedschaften von $($SourceUser.SamAccountName) zu $($newUser.SamAccountName)."
                if ($PSCmdlet.ShouldProcess($newUser.SamAccountName, "Gruppenmitgliedschaften hinzufügen ($($groupsToCopy.Count) Gruppen)")) {
                    Add-ADPrincipalGroupMembership -Identity $newUser -MemberOf $groupsToCopy -ErrorAction Stop
                    Write-Log -Level Info -Message "Gruppenmitgliedschaften erfolgreich kopiert."
                    # Optional: Update Report
                    # Add-UserReportEntry -SamAccountName $newUser.SamAccountName -Status "Gruppen kopiert" -Detail "$($groupsToCopy.Count) Gruppen"
                } else {
                     Write-Log -Level Info -Message "Kopieren der Gruppenmitgliedschaften übersprungen (ShouldProcess)."
                }
            } else {
                Write-Log -Level Info -Message "Quellbenutzer $($SourceUser.SamAccountName) hat keine (relevanten) Gruppenmitgliedschaften zum Kopieren."
            }
        }
        catch {
            $msg = "Fehler beim Kopieren der Gruppenmitgliedschaften für '$($newUser.SamAccountName)': $_. Der Benutzer wurde erstellt, aber Gruppen fehlen möglicherweise."
            Write-Log -Level Warning -Message $msg
            Add-UserReportEntry -SamAccountName $newUser.SamAccountName -Status "Warnung" -Detail "Fehler beim Kopieren der Gruppen: $_"
            # Nicht abbrechen, Benutzer existiert ja schon
        }

        Write-Log -Level Info -Message "Kopiervorgang für $TargetSamAccountName abgeschlossen."
        return $newUser
    }

     # Funktion zum Erstellen eines Benutzers aus Daten (CSV-Zeile/Hashtable)
     function New-ADUserFromData {
         [CmdletBinding(SupportsShouldProcess = $true)]
         param(
             [Parameter(Mandatory = $true)]
             [hashtable]$UserData, # Enthält alle Infos aus CSV-Zeile

             [Parameter(Mandatory = $false)]
             [Microsoft.ActiveDirectory.Management.ADUser]$TemplateUser, # Optional: Template User Objekt (ReferenceUser)

             [Parameter(Mandatory = $false)]
             [System.Security.SecureString]$GlobalDefaultPassword, # Optional: Fallback Passwort

             [Parameter(Mandatory = $false)]
             [string]$GlobalTargetOU # Optional: Fallback OU
         )

         # Versuche SamAccountName zu bekommen, bevor geloggt wird
         $sam = $UserData.SamAccountName
         if (-not $sam) {
             # Wenn SamAccountName fehlt, können wir nicht viel tun
             $msg = "Fehlender Wert für 'SamAccountName' in den Daten. Überspringe Eintrag."
             Write-Log -Level Error -Message $msg
             Add-UserReportEntry -SamAccountName "(Unbekannt)" -Status "Fehler" -Detail $msg
             return $null
         }

         Write-Log -Level Info -Message "Beginne Verarbeitung zur Erstellung von Benutzer '$sam'."

         # --- Validierung der Pflichtfelder ---
         if (-not $UserData.GivenName) {
             $msg = "Fehlender Wert für 'GivenName' (Vorname) für '$sam'. Überspringe Eintrag."
             Write-Log -Level Error -Message $msg
             Add-UserReportEntry -SamAccountName $sam -Status "Fehler" -Detail $msg
             return $null
         }
         if (-not $UserData.Surname) {
             $msg = "Fehlender Wert für 'Surname' (Nachname) für '$sam'. Überspringe Eintrag."
             Write-Log -Level Error -Message $msg
             Add-UserReportEntry -SamAccountName $sam -Status "Fehler" -Detail $msg
             return $null
         }

         # --- Ziel-OU bestimmen (Priorität: Global Parameter > CSV > Template > Fehler) ---
         $finalOU = $null
         if ($GlobalTargetOU) {
             $finalOU = $GlobalTargetOU
             Write-Verbose "Verwende globale Ziel-OU '$finalOU' für '$sam'."
         } elseif ($UserData.ContainsKey('TargetOU') -and $UserData.TargetOU) {
             $finalOU = $UserData.TargetOU
             Write-Verbose "Verwende Ziel-OU aus Datenquelle '$finalOU' für '$sam'."
         } elseif ($TemplateUser) {
             $finalOU = ($TemplateUser.DistinguishedName -split ',', 2)[1]
             Write-Verbose "Verwende Ziel-OU vom Template-Benutzer '$finalOU' für '$sam'."
         } else {
             $msg = "Keine Ziel-OU für Benutzer '$sam' gefunden (weder in CSV, noch als Parameter, noch durch Template). Überspringe Eintrag."
             Write-Log -Level Error -Message $msg
             Add-UserReportEntry -SamAccountName $sam -Status "Fehler" -Detail $msg
             return $null
         }

         # --- Passwort bestimmen (Priorität: CSV > Global Parameter > Generieren) ---
         $finalPassword = $null
         $changePwdAtLogon = $true
         if ($UserData.ContainsKey('Password') -and $UserData.Password) {
             Write-Log -Level Warning -Message "Verwende Passwort aus Datenquelle für '$sam'. ACHTUNG: Klartextpasswörter sind ein Sicherheitsrisiko!"
             try {
                 $finalPassword = ConvertTo-SecureString $UserData.Password -AsPlainText -Force -ErrorAction Stop
             } catch {
                 $msg = "Fehler beim Konvertieren des Passworts aus der Datenquelle für '$sam': $_. Überspringe Eintrag."
                 Write-Log -Level Error -Message $msg
                 Add-UserReportEntry -SamAccountName $sam -Status "Fehler" -Detail $msg
                 return $null
             }
         } elseif ($GlobalDefaultPassword) {
             Write-Verbose "Verwende globales Standardpasswort für '$sam'."
             $finalPassword = $GlobalDefaultPassword
         } else {
             Write-Log -Level Info -Message "Generiere zufälliges Passwort für '$sam', da keines angegeben wurde."
             # Generiere sicheres, zufälliges Passwort
             try {
                 # Komplexeres Beispiel: Mind. 1 Groß, 1 Klein, 1 Zahl, 1 Sonderzeichen, Länge 14
                 $pwdChars = @()
                 $pwdChars += 65..90 | Get-Random # Großbuchstabe
                 $pwdChars += 97..122 | Get-Random # Kleinbuchstabe
                 $pwdChars += 48..57 | Get-Random # Zahl
                 $pwdChars += 33, 35, 36, 37, 38, 42, 64, 95 | Get-Random # Sonderzeichen !#$%&*@_
                 # Restliche Zeichen auffüllen (insgesamt 14)
                 $allChars = (48..57) + (65..90) + (97..122) + 33, 35, 36, 37, 38, 42, 64, 95
                 $pwdChars += $allChars | Get-Random -Count (14 - $pwdChars.Count)
                 # Mischen
                 $randomPassword = -join ($pwdChars | Get-Random -Count $pwdChars.Count | % {[char]$_})

                 $finalPassword = ConvertTo-SecureString $randomPassword -AsPlainText -Force -ErrorAction Stop
                 Write-Log -Level Info -Message "Zufälliges Passwort für '$sam' generiert. Benutzer MUSS es bei der ersten Anmeldung ändern."
             } catch {
                 $msg = "Fehler beim Generieren/Konvertieren des zufälligen Passworts für '$sam': $_. Überspringe Eintrag."
                 Write-Log -Level Error -Message $msg
                 Add-UserReportEntry -SamAccountName $sam -Status "Fehler" -Detail $msg
                 return $null
             }
         }

         # --- Prüfen ob Zielbenutzer existiert ---
         try {
             if (Get-ADUser -Filter "SamAccountName -eq '$sam'" -ErrorAction SilentlyContinue) {
                 $msg = "Benutzer '$sam' existiert bereits im AD. Überspringe Eintrag."
                 Write-Log -Level Error -Message $msg
                 Add-UserReportEntry -SamAccountName $sam -Status "Fehler" -Detail $msg
                 return $null
             }
         } catch { Write-Log -Level Warning -Message "Fehler beim Prüfen, ob Benutzer '$sam' existiert: $_." }


        # --- Prüfen ob Ziel-OU existiert ---
        try {
            if (-not (Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$finalOU'" -ErrorAction Stop)) {
                 $msg = "Die Ziel-OU '$finalOU' für '$sam' existiert nicht. Überspringe Eintrag."
                 Write-Log -Level Error -Message $msg
                 Add-UserReportEntry -SamAccountName $sam -Status "Fehler" -Detail $msg
                 return $null
            }
            Write-Verbose "Ziel-OU '$finalOU' für '$sam' ist gültig."
        } catch {
            $msg = "Fehler beim Überprüfen der Ziel-OU '$finalOU' für '$sam': $_. Überspringe Eintrag."
            Write-Log -Level Error -Message $msg
            Add-UserReportEntry -SamAccountName $sam -Status "Fehler" -Detail $msg
            return $null
        }

         # --- Benutzerparameter zusammenstellen ---
         $newUserParams = @{
             SamAccountName        = $sam
             Name                  = "$($UserData.GivenName) $($UserData.Surname)" # Standard Name
             GivenName             = $UserData.GivenName
             Surname               = $UserData.Surname
             DisplayName           = "$($UserData.GivenName) $($UserData.Surname)" # Standard DisplayName
             UserPrincipalName     = "$sam@$($env:USERDNSDOMAIN)" # Anpassen falls nötig
             Path                  = $finalOU
             AccountPassword       = $finalPassword
             ChangePasswordAtLogon = $changePwdAtLogon
             Enabled               = $true # Standardmäßig aktivieren
         }

         # Enabled-Status aus Datenquelle übernehmen, falls vorhanden
         if ($UserData.ContainsKey('Enabled')) {
             try {
                 $newUserParams.Enabled = [bool]::Parse($UserData.Enabled) # Sicherere Konvertierung
                 Write-Verbose "Setze 'Enabled' für '$sam' auf '$($newUserParams.Enabled)' basierend auf Datenquelle."
             } catch {
                 Write-Log -Level Warning -Message "Konnte 'Enabled'-Wert '$($UserData.Enabled)' für '$sam' nicht in Boolean konvertieren. Verwende Standard ($true)."
             }
         }

         # Weitere Attribute aus Datenquelle oder Template übernehmen
         $attributesToCheck = @(
             'Description', 'Office', 'Department', 'Title', 'Company',
             'EmailAddress', 'StreetAddress', 'City', 'State', 'PostalCode', 'Country', 'OfficePhone'
             # Fügen Sie hier weitere AD-Attribute hinzu, die unterstützt werden sollen
         )

         foreach ($attr in $attributesToCheck) {
             if ($UserData.ContainsKey($attr) -and $UserData.$attr) {
                 $newUserParams[$attr] = $UserData.$attr
                 Write-Verbose "Setze Attribut '$attr' für '$sam' aus Datenquelle."
             } elseif ($TemplateUser -and $TemplateUser.$attr) {
                 # Nur übernehmen, wenn Attribut im Template existiert und nicht leer ist
                 if ($TemplateUser.PSObject.Properties[$attr] -ne $null -and $TemplateUser.$attr -ne '') {
                    $newUserParams[$attr] = $TemplateUser.$attr
                    Write-Verbose "Setze Attribut '$attr' für '$sam' vom Template-Benutzer."
                 }
             }
         }

         # --- Benutzer erstellen ---
         $newUser = $null
         if ($PSCmdlet.ShouldProcess($sam, "Neuen AD-Benutzer erstellen")) {
             try {
                 Write-Log -Level Info -Message "Erstelle Benutzer '$sam' in OU '$finalOU'."
                 $newUser = New-ADUser @newUserParams -PassThru -ErrorAction Stop
                 Write-Log -Level Info -Message "Benutzer '$($newUser.SamAccountName)' erfolgreich erstellt (SID: $($newUser.SID.Value))."
                 Add-UserReportEntry -SamAccountName $newUser.SamAccountName -Status "Erstellt" -Detail "In OU '$finalOU'"
             } catch {
                 $msg = "Fehler beim Erstellen des Benutzers '$sam': $_"
                 Write-Log -Level Error -Message $msg
                 Add-UserReportEntry -SamAccountName $sam -Status "Fehler" -Detail $msg
                 return $null # Abbruch für diesen Benutzer
             }
         } else {
             $msg = "Erstellung von '$sam' übersprungen (ShouldProcess)."
             Write-Log -Level Info -Message $msg
             Add-UserReportEntry -SamAccountName $sam -Status "Übersprungen" -Detail $msg
             return $null # Abbruch für diesen Benutzer
         }

         # --- Gruppen vom Template übernehmen ---
         if ($TemplateUser) {
             try {
                 $templateGroups = Get-ADPrincipalGroupMembership -Identity $TemplateUser -ErrorAction Stop
                 $groupsToCopy = $templateGroups | Where-Object {$_.Name -ne "Domain Users"} # Filter

                 if ($groupsToCopy) {
                     Write-Log -Level Info -Message "Füge Benutzer '$($newUser.SamAccountName)' zu $($groupsToCopy.Count) Gruppen hinzu (basierend auf Template '$($TemplateUser.SamAccountName)')."
                     if ($PSCmdlet.ShouldProcess($newUser.SamAccountName, "Gruppenmitgliedschaften hinzufügen ($($groupsToCopy.Count) Gruppen)")) {
                         Add-ADPrincipalGroupMembership -Identity $newUser -MemberOf $groupsToCopy -ErrorAction Stop
                         Write-Log -Level Info -Message "Gruppenmitgliedschaften für '$($newUser.SamAccountName)' erfolgreich hinzugefügt."
                         # Optional: Update Report
                         # Add-UserReportEntry -SamAccountName $newUser.SamAccountName -Status "Gruppen hinzugefügt" -Detail "$($groupsToCopy.Count) Gruppen von $($TemplateUser.SamAccountName)"
                     } else {
                         Write-Log -Level Info -Message "Hinzufügen der Gruppenmitgliedschaften übersprungen (ShouldProcess)."
                     }
                 } else {
                     Write-Log -Level Info -Message "Template-Benutzer $($TemplateUser.SamAccountName) hat keine (relevanten) Gruppenmitgliedschaften zum Hinzufügen."
                 }
             } catch {
                 $msg = "Fehler beim Hinzufügen der Gruppenmitgliedschaften vom Template für '$($newUser.SamAccountName)': $_. Der Benutzer wurde erstellt, aber Gruppen fehlen möglicherweise."
                 Write-Log -Level Warning -Message $msg
                 Add-UserReportEntry -SamAccountName $newUser.SamAccountName -Status "Warnung" -Detail "Fehler beim Hinzufügen der Gruppen: $_"
             }
         }

         Write-Log -Level Info -Message "Verarbeitung für Benutzer '$sam' abgeschlossen."
         return $newUser
     }

    Write-Verbose "Initialisierung abgeschlossen. Wechsle zur Prozess-Phase."
} # End Begin Block

process {
    Write-Verbose "Beginne Prozess-Phase. Ausgewählter Modus: $($PSCmdlet.ParameterSetName)"

    switch ($PSCmdlet.ParameterSetName) {
        # --- Modus: CopySingleUser ---
        'CopySingleUser' {
            Write-Log -Level Info -Message "Starte Modus: CopySingleUser"

            # 1. Referenzbenutzer (Quelle) validieren
            $referenceUserObject = $null
            try {
                # Verwende den konsolidierten Parameter $ReferenceUserSamAccountName
                $referenceUserObject = Get-ADUser -Identity $ReferenceUserSamAccountName -Properties * -ErrorAction Stop # Lade alle Properties für die Kopie
                Write-Log -Level Info -Message "Referenzbenutzer (Quelle) '$($referenceUserObject.SamAccountName)' gefunden."
            } catch {
                $msg = "Referenzbenutzer (Quelle) '$ReferenceUserSamAccountName' konnte nicht gefunden werden: $_"
                Write-Log -Level Error -Message $msg
                Add-UserReportEntry -SamAccountName $ReferenceUserSamAccountName -Status "Fehler" -Detail $msg
                return # Abbruch des Modus
            }

            # 2. Zielbenutzernamen abfragen, wenn nicht angegeben
            if (-not $TargetUserSamAccountName) {
                try {
                    $TargetUserSamAccountName = Read-Host -Prompt "Bitte geben Sie den gewünschten SAMAccountName für den neuen Benutzer ein"
                    if (-not $TargetUserSamAccountName) { throw "Eingabe darf nicht leer sein."}
                } catch {
                     $msg = "Ungültige Eingabe für Zielbenutzernamen: $_"
                     Write-Log -Level Error -Message $msg
                     Add-UserReportEntry -SamAccountName "(Unbekannt)" -Status "Fehler" -Detail $msg
                     return
                }
            }

            # 3. Passwort abfragen, wenn nicht angegeben
            if (-not $TargetUserPassword) {
                 try {
                     $TargetUserPassword = Read-Host -Prompt "Bitte geben Sie das initiale Passwort für '$TargetUserSamAccountName' ein" -AsSecureString
                     if ($TargetUserPassword.Length -eq 0) { throw "Passwort darf nicht leer sein."} # Einfache Prüfung
                 } catch {
                     $msg = "Ungültige Eingabe für Passwort: $_"
                     Write-Log -Level Error -Message $msg
                     Add-UserReportEntry -SamAccountName $TargetUserSamAccountName -Status "Fehler" -Detail $msg
                     return
                 }
            }

            # 4. Kopierfunktion aufrufen
            $newUser = Copy-ADUserAdvanced -SourceUser $referenceUserObject `
                                          -TargetSamAccountName $TargetUserSamAccountName `
                                          -Password $TargetUserPassword `
                                          -DestinationOU $TargetOU `
                                          -OverwriteTarget:$Force `
                                          -Verbose:$VerbosePreference.ToString() -WarningAction $WarningPreference # Explizit übergeben

            if ($newUser) {
                Write-Log -Level Info -Message "Benutzer '$($newUser.SamAccountName)' erfolgreich kopiert."
                # Report-Eintrag wird bereits in Copy-ADUserAdvanced hinzugefügt
            } else {
                Write-Log -Level Error -Message "Fehler beim Kopieren des Benutzers '$ReferenceUserSamAccountName'."
                # Fehler und Report-Eintrag wurden bereits in Copy-ADUserAdvanced geloggt/hinzugefügt
            }
        } # End CopySingleUser

        # --- Modus: CreateUsersFromCSV ---
        'CreateUsersFromCSV' {
            Write-Log -Level Info -Message "Starte Modus: CreateUsersFromCSV"
            $SuccessCount = 0
            $FailCount = 0

            # 1. CSV-Datei einlesen
            $usersData = @()
            try {
                Write-Log -Level Info -Message "Lese CSV-Datei: $CsvPath"
                $usersData = Import-Csv -Path $CsvPath -Delimiter ';' -Encoding UTF8 -ErrorAction Stop
                Write-Log -Level Info -Message "$($usersData.Count) Einträge in CSV-Datei gefunden."
            } catch {
                $msg = "Fehler beim Lesen der CSV-Datei '$CsvPath': $_"
                Write-Log -Level Error -Message $msg
                Add-UserReportEntry -SamAccountName "(CSV)" -Status "Fehler" -Detail $msg
                return # Abbruch des Modus
            }

            if ($usersData.Count -eq 0) {
                 $msg = "CSV-Datei '$CsvPath' ist leer oder enthält keine Daten."
                 Write-Log -Level Warning -Message $msg
                 Add-UserReportEntry -SamAccountName "(CSV)" -Status "Warnung" -Detail $msg
                 return
            }

            # 2. Referenzbenutzer (Template) laden (wenn angegeben)
            $referenceUserObject = $null
            # Verwende den konsolidierten Parameter $ReferenceUserSamAccountName
            if ($ReferenceUserSamAccountName) {
                try {
                    $referenceUserObject = Get-ADUser -Identity $ReferenceUserSamAccountName -Properties * -ErrorAction Stop # Alle Properties laden
                    Write-Log -Level Info -Message "Referenzbenutzer (Template) '$($referenceUserObject.SamAccountName)' gefunden und geladen."
                } catch {
                    Write-Log -Level Warning -Message "Referenzbenutzer (Template) '$ReferenceUserSamAccountName' konnte nicht gefunden werden: $_. Gruppen und Attribute werden nicht vom Template übernommen."
                    # Nicht abbrechen, aber weitermachen ohne Template
                }
            } else {
                 Write-Log -Level Info -Message "Kein Referenzbenutzer (Template) angegeben (-ReferenceUserSamAccountName). Attribute/Gruppen werden nur aus CSV übernommen."
            }

            # 3. Jeden Eintrag in der CSV verarbeiten
            Write-Log -Level Info -Message "Beginne Verarbeitung der CSV-Einträge..."
            foreach ($userRow in $usersData) {
                $userDataHash = @{}
                # Konvertiere PSObject aus Import-Csv in Hashtable für einfachere Handhabung
                $userRow.PSObject.Properties | ForEach-Object { $userDataHash[$_.Name] = $_.Value }

                $newUser = New-ADUserFromData -UserData $userDataHash `
                                              -TemplateUser $referenceUserObject `
                                              -GlobalDefaultPassword $DefaultPassword `
                                              -GlobalTargetOU $TargetOU `
                                              -Verbose:$VerbosePreference.ToString() -WarningAction $WarningPreference

                if ($newUser) {
                    $SuccessCount++
                } else {
                    $FailCount++
                    # Fehler und Report-Eintrag wurden bereits in New-ADUserFromData geloggt/hinzugefügt
                }
            } # End foreach userRow

            Write-Log -Level Info -Message "CSV-Verarbeitung abgeschlossen. Erfolgreich erstellt: $SuccessCount, Fehlgeschlagen/Übersprungen: $FailCount."

        } # End CreateUsersFromCSV

        default {
            # Sollte nicht passieren bei korrekter Parameternutzung
            Write-Log -Level Error -Message "Unbekannter oder keiner der Hauptmodi wurde ausgewählt. Verwenden Sie -CopySingleUser oder -CreateUsersFromCSV."
            Add-UserReportEntry -SamAccountName "(Skript)" -Status "Fehler" -Detail "Ungültiger Modus"
        }
    } # End Switch ParameterSetName

    Write-Verbose "Prozess-Phase abgeschlossen."

} # End Process Block

end {
    $scriptEndTime = Get-Date
    $duration = New-TimeSpan -Start $scriptStartTime -End $scriptEndTime
    Write-Verbose "Beginne End-Phase."

    # Benutzerbericht schreiben, wenn aktiviert und Daten vorhanden
    if (-not $NoUserReport -and $global:userReportData.Count -gt 0) {
        Write-Log -Level Info -Message "Schreibe Benutzerbericht nach '$($global:fullUserReportPath)'..."
        if ($PSCmdlet.ShouldProcess($global:fullUserReportPath, "Benutzerbericht exportieren")) {
            try {
                $global:userReportData | Export-Csv -Path $global:fullUserReportPath -Delimiter ';' -NoTypeInformation -Encoding UTF8 -Force -ErrorAction Stop
                Write-Log -Level Info -Message "Benutzerbericht erfolgreich geschrieben."
            } catch {
                Write-Log -Level Error -Message "Fehler beim Schreiben des Benutzerberichts nach '$($global:fullUserReportPath)': $_"
            }
        } else {
             Write-Log -Level Info -Message "Schreiben des Benutzerberichts übersprungen (ShouldProcess)."
        }
    } elseif (-not $NoUserReport) {
         Write-Log -Level Info -Message "Keine Daten für Benutzerbericht vorhanden."
    }

    Write-Log -Level Info -Message "Skriptausführung beendet. Gesamtdauer: $($duration.ToString('g'))"
    Write-Log -Level Info -Message "Log-Datei: $fullLogPath"
    Write-Verbose "End-Phase abgeschlossen."
} # End End Block
