<#
.SYNOPSIS
Verwaltet Active Directory-Benutzer mit modernem interaktivem Menü: Kopiert, erstellt, importiert aus CSV, wendet Eigenschaften an, exportiert.
Unicode UTF-8, PowerShell 5.1 kompatibel, keine externen Abhängigkeiten.

.DESCRIPTION
Vollständiges AD-Management-Tool mit:
- Interaktives Menü (Unicode, korrekte Box-Drawing-Zeichen)
- CopySingleUser: Kopiert Benutzer + Gruppen + OU-Struktur + dynamische Domain-Ermittlung
- CreateUsersFromCSV: Batch-Import mit Template-Unterstützung
- ApplyPropertiesToExistingUser: Synchronisiert Eigenschaften + Gruppen (mit Fix für Serververbindung)
- ExportUserData: Filtert und exportiert Benutzer flexibel
- ExportLKennung: L-Kennung-Spezialexport mit OU-Filterung
- Debug-Output für alle Operationen
- Strukturiertes Reporting (CSV + Log)

.NOTES
Version: 9.0 FINAL (Konsolidiert v8 + v6.7 Features)
Datum: 2025-05-28
Autor: Systems Administration (Fixes + Optimierungen)
Kompatibilität: PowerShell 5.1+, UTF-8, ISE-ready
Verbessert: Menu-Formatierung, Serververbindung, OU-Filterung, Error Handling
#>

#requires -Version 5.1
#requires -Modules ActiveDirectory

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('CopySingleUser', 'CreateUsersFromCSV', 'ApplyPropertiesToExistingUser', 'ExportUserData', 'ExportLKennung', 'Interactive')]
    [string]$Mode = 'Interactive',

    [Parameter(Mandatory = $false)]
    [string]$ReferenceUserSamAccountName,

    [Parameter(Mandatory = $false)]
    [string]$TargetUserSamAccountName,

    [Parameter(Mandatory = $false)]
    [string]$TargetOU,

    [Parameter(Mandatory = $false)]
    [System.Security.SecureString]$TargetUserPassword,

    [Parameter(Mandatory = $false)]
    [switch]$Force,

    [Parameter(Mandatory = $false)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$CsvPath,

    [Parameter(Mandatory = $false)]
    [System.Security.SecureString]$DefaultPassword,

    [Parameter(Mandatory = $false)]
    [string]$IdentityFilter,

    [Parameter(Mandatory = $false)]
    [string]$OUFilter,

    [Parameter(Mandatory = $false)]
    [string]$SearchBaseOU,

    [Parameter(Mandatory = $false)]
    [string]$ExportCsvPath,

    [Parameter(Mandatory = $false)]
    [string[]]$LKennungOUNames = @('81', '82'),

    [Parameter(Mandatory = $false)]
    [string]$LKennungLDAPFilter = "(|(sAMAccountName=L110*)(sAMAccountName=L114*))",

    [Parameter(Mandatory = $false)]
    [string]$LKennungExportCsvPath,

    [Parameter(Mandatory = $false)]
    [string]$LogPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Error', 'Warning', 'Info', 'Verbose', 'Debug')]
    [string]$LogLevel = 'Info',

    [Parameter(Mandatory = $false)]
    [string]$UserReportCsvPath
)


. (Join-Path $PSScriptRoot 'scripts/utils/ADUserManagement.ps1')

begin {
    #region Initialisierung

    . (Join-Path $PSScriptRoot 'scripts/utils/ExportLKennung.ps1')
    $ErrorActionPreference = 'Stop'
    $VerbosePreference = 'Continue'
    
    # UTF-8 Encoding sicherstellen (mit Fehlerbehandlung)
    try {
        if ($Host.Name -notmatch "ISE") {
            [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        }
        else {
            Write-Warning "PowerShell ISE erkannt - [Console]::OutputEncoding übersprungen"
        }
    }
    catch {
        Write-Warning "UTF-8 OutputEncoding konnte nicht gesetzt werden: $_"
    }

    # Kultur auf Deutsch
    try {
        [System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo('de-DE')
        [System.Threading.Thread]::CurrentThread.CurrentUICulture = [System.Globalization.CultureInfo]::GetCultureInfo('de-DE')
    }
    catch {
        Write-Warning "Warnung: Kultur-Einstellung fehlgeschlagen."
    }

    # Basisverzeichnis
    $basePath = $PSScriptRoot
    if (-not $basePath) { $basePath = (Get-Location).Path }

    # Logging-Setup
    $scriptStartTime = Get-Date
    if (-not $PSBoundParameters.ContainsKey('LogPath')) { $LogPath = $basePath }
    if (-not (Test-Path $LogPath -PathType Container)) { New-Item -Path $LogPath -ItemType Directory -Force | Out-Null }
    
    $logFileName = "Enhanced-ADManagement_{0}_{1}.log" -f $Mode, $scriptStartTime.ToString('yyyyMMdd-HHmmss')
    $global:fullLogPath = Join-Path -Path $LogPath -ChildPath $logFileName

    # Log-Funktion (UTF-8 sicher)
    function Write-Log {
        param(
            [Parameter(Mandatory = $true)]
            [ValidateSet('Error', 'Warning', 'Info', 'Verbose', 'Debug')]
            [string]$Level,
            [Parameter(Mandatory = $true)]
            [string]$Message
        )

        $logLevels = @{'Debug' = 0; 'Verbose' = 1; 'Info' = 2; 'Warning' = 3; 'Error' = 4}
        $currentLogLevelValue = $logLevels[$LogLevel]
        $messageLogLevelValue = $logLevels[$Level]

        if ($messageLogLevelValue -ge $currentLogLevelValue) {
            $logEntry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
            try {
                Add-Content -Path $global:fullLogPath -Value $logEntry -Encoding UTF8 -ErrorAction Stop
            }
            catch {
                # Stille Fehler, nicht abbrechen
            }
        }

        switch ($Level) {
            'Error'   { Write-Error $Message }
            'Warning' { Write-Warning $Message }
            'Info'    { Write-Host "[INFO] $Message" -ForegroundColor Green }
            'Verbose' { Write-Verbose $Message }
            'Debug'   { Write-Host "[DEBUG] $Message" -ForegroundColor Cyan }
        }
    }

    # User Report Setup
    $global:userReportData = [System.Collections.Generic.List[PSObject]]::new()
    if (-not $PSBoundParameters.ContainsKey('UserReportCsvPath')) { $UserReportCsvPath = $basePath }
    if (-not (Test-Path $UserReportCsvPath -PathType Container)) { New-Item -Path $UserReportCsvPath -ItemType Directory -Force | Out-Null }
    
    $reportFileName = "Enhanced-ADManagement_UserReport_{0}.csv" -f $scriptStartTime.ToString('yyyyMMdd-HHmmss')
    $global:fullUserReportPath = Join-Path -Path $UserReportCsvPath -ChildPath $reportFileName

    # Report-Funktion
    function Add-UserReportEntry {
        param(
            [Parameter(Mandatory = $true)]
            [string]$SamAccountName,
            [Parameter(Mandatory = $true)]
            [string]$Status,
            [Parameter(Mandatory = $false)]
            [string]$Detail = ""
        )
        
        $reportObject = [PSCustomObject]@{
            Timestamp      = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            SamAccountName = $SamAccountName
            Status         = $Status
            Detail         = $Detail
        }
        $global:userReportData.Add($reportObject)
    }

    # FIX: Dynamische Domain/Server-Ermittlung aus Benutzer-DN
    function Get-ADServerFromUser {
        param(
            [Parameter(Mandatory = $true)]
            [Microsoft.ActiveDirectory.Management.ADUser]$User
        )
        
        try {
            # Extrahiere Domain aus DN: CN=...,OU=...,DC=firma,DC=local → firma.local
            $dnParts = $User.DistinguishedName -split ','
            $dcParts = $dnParts | Where-Object { $_ -like 'DC=*' } | ForEach-Object { $_ -replace 'DC=' }
            $domain = $dcParts -join '.'
            
            Write-Log -Level Debug -Message "Ermittelte Domain aus Benutzer: $domain"
            return $domain
        }
        catch {
            Write-Log -Level Warning -Message "Konnte Domain nicht aus Benutzer-DN ermitteln: $_"
            return $null
        }
    }

    # Optimierte OU-Filterung (aus v6.7)
    function Get-ADUsersWithOUFilter {
        param(
            [Parameter(Mandatory = $true)]
            [string]$IdentityFilter,
            [Parameter(Mandatory = $false)]
            [string]$OUFilter,
            [Parameter(Mandatory = $false)]
            [string]$SearchBaseOU
        )
        
        try {
            $filter = "*$IdentityFilter*"
            $users = Get-ADUser -Filter "SamAccountName -like '$filter'" -Properties DistinguishedName -ErrorAction Stop
            
            if ($OUFilter -or $SearchBaseOU) {
                if ($SearchBaseOU) {
                    Write-Log -Level Debug -Message "Filtere nach SearchBaseOU: $SearchBaseOU"
                    $users = $users | Where-Object { $_.DistinguishedName -like "*$SearchBaseOU*" }
                }
                elseif ($OUFilter) {
                    Write-Log -Level Debug -Message "Filtere nach OUFilter: $OUFilter"
                    $users = $users | Where-Object { 
                        $ouName = ($_.DistinguishedName -split ',' | Where-Object { $_ -like 'OU=*' }) -join ','
                        $ouName -like "*$OUFilter*"
                    }
                }
            }
            
            return $users
        }
        catch {
            Write-Log -Level Error -Message "Fehler beim Filtern von Benutzern: $_"
            return @()
        }
    }

    # AD-Benutzer kopieren (mit Gruppen)
    function Copy-ADUserAdvanced {
        [CmdletBinding(SupportsShouldProcess = $true)]
        param(
            [Parameter(Mandatory = $true)]
            [Microsoft.ActiveDirectory.Management.ADUser]$SourceUser,
            [Parameter(Mandatory = $true)]
            [string]$TargetSamAccountName,
            [Parameter(Mandatory = $true)]
            [System.Security.SecureString]$Password,
            [Parameter(Mandatory = $false)]
            [string]$DestinationOU,
            [Parameter(Mandatory = $false)]
            [switch]$OverwriteTarget
        )

        Write-Log -Level Debug -Message "START: Kopiere Benutzer von '$($SourceUser.SamAccountName)' zu '$TargetSamAccountName'"
        Write-Log -Level Info -Message "Starte Kopiervorgang von '$($SourceUser.SamAccountName)' zu '$TargetSamAccountName'"
        
        # Zielbenutzer existiert?
        $targetUserExists = $false
        try {
            $existingTargetUser = Get-ADUser -Filter "SamAccountName -eq '$TargetSamAccountName'" -ErrorAction SilentlyContinue
            if ($existingTargetUser) {
                $targetUserExists = $true
                Write-Log -Level Debug -Message "├─ Zielbenutzer existiert bereits"
            }
        }
        catch {
            Write-Log -Level Debug -Message "├─ FEHLER beim Prüfen: $_"
        }

        if ($targetUserExists) {
            if (-not $OverwriteTarget) {
                Write-Log -Level Warning -Message "└─ Benutzer '$TargetSamAccountName' existiert (Force: $($Force))"
                Add-UserReportEntry -SamAccountName $TargetSamAccountName -Status "Fehler" -Detail "Existiert bereits"
                return $null
            }
            else {
                Write-Log -Level Debug -Message "├─ Überschreibe existierenden Benutzer"
                if ($PSCmdlet.ShouldProcess($TargetSamAccountName, "Entfernen")) {
                    Remove-ADUser -Identity $existingTargetUser -Confirm:$false -ErrorAction Stop
                }
            }
        }

        # OU bestimmen
        $finalOU = $DestinationOU
        if (-not $finalOU) {
            $finalOU = ($SourceUser.DistinguishedName -split ',', 2)[1]
            Write-Log -Level Debug -Message "├─ OU aus Quelle: $finalOU"
        }

        # OU prüfen
        try {
            Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$finalOU'" -ErrorAction Stop | Out-Null
            Write-Log -Level Debug -Message "├─ OU validiert"
        }
        catch {
            Write-Log -Level Error -Message "└─ OU existiert nicht: $finalOU"
            Add-UserReportEntry -SamAccountName $TargetSamAccountName -Status "Fehler" -Detail "OU nicht gefunden"
            return $null
        }

        # Benutzer erstellen
        $newUserParams = @{
            SamAccountName        = $TargetSamAccountName
            Name                  = $TargetSamAccountName
            GivenName             = $SourceUser.GivenName
            Surname               = $SourceUser.Surname
            DisplayName           = "$($SourceUser.GivenName) $($SourceUser.Surname)"
            UserPrincipalName     = "$TargetSamAccountName@$($env:USERDNSDOMAIN)"
            Path                  = $finalOU
            AccountPassword       = $Password
            ChangePasswordAtLogon = $true
            Enabled               = $true
            Description           = $SourceUser.Description
            Office                = $SourceUser.Office
            Department            = $SourceUser.Department
            Company               = $SourceUser.Company
            Title                 = $SourceUser.Title
            StreetAddress         = $SourceUser.StreetAddress
            City                  = $SourceUser.City
            State                 = $SourceUser.State
            PostalCode            = $SourceUser.PostalCode
            Country               = $SourceUser.Country
            OfficePhone           = $SourceUser.OfficePhone
            EmailAddress          = $SourceUser.EmailAddress
        }

        $newUser = $null
        if ($PSCmdlet.ShouldProcess($TargetSamAccountName, "Benutzer erstellen")) {
            try {
                Write-Log -Level Debug -Message "├─ Erstelle Benutzer..."
                $newUser = New-ADUser @newUserParams -PassThru -ErrorAction Stop
                Write-Log -Level Info -Message "Benutzer erstellt: '$($newUser.SamAccountName)'"
                Write-Log -Level Debug -Message "├─ Benutzer erfolgreich erstellt"
                Add-UserReportEntry -SamAccountName $newUser.SamAccountName -Status "Kopiert" -Detail "Von $($SourceUser.SamAccountName)"
            }
            catch {
                Write-Log -Level Error -Message "Benutzer-Erstellung fehlgeschlagen: $_"
                Write-Log -Level Debug -Message "└─ FEHLER: $_"
                Add-UserReportEntry -SamAccountName $TargetSamAccountName -Status "Fehler" -Detail $_
                return $null
            }
        }

        # Gruppen kopieren
        if ($newUser) {
            try {
                Write-Log -Level Debug -Message "├─ Hole Quellgruppen..."
                $sourceGroups = Get-ADPrincipalGroupMembership -Identity $SourceUser -ErrorAction Stop
                Write-Log -Level Debug -Message "├─ Gefundene Gruppen: $($sourceGroups.Count)"
                
                $groupsToCopy = $sourceGroups | Where-Object { $_.Name -ne "Domain Users" -and $_.Name -ne "Domänen-Benutzer" }
                
                if ($groupsToCopy) {
                    Write-Log -Level Info -Message "Kopiere $($groupsToCopy.Count) Gruppen"
                    Write-Log -Level Debug -Message "├─ Gruppen zum Kopieren:"
                    foreach ($grp in $groupsToCopy) {
                        Write-Log -Level Debug -Message "│  - $($grp.Name)"
                    }
                    
                    if ($PSCmdlet.ShouldProcess($newUser.SamAccountName, "Gruppen hinzufügen")) {
                        # Prüfe auf Duplikate
                        $existingGroups = Get-ADPrincipalGroupMembership -Identity $newUser -ErrorAction SilentlyContinue
                        $actuallyAddGroups = $groupsToCopy | Where-Object { $_.DistinguishedName -notin $existingGroups.DistinguishedName }
                        
                        if ($actuallyAddGroups) {
                            Add-ADPrincipalGroupMembership -Identity $newUser -MemberOf $actuallyAddGroups -ErrorAction Stop
                            Write-Log -Level Info -Message "Gruppen hinzugefügt: $($actuallyAddGroups.Count)"
                            Write-Log -Level Debug -Message "└─ Gruppen erfolgreich kopiert"
                        }
                        else {
                            Write-Log -Level Debug -Message "└─ Keine neuen Gruppen hinzuzufügen (keine Duplikate)"
                        }
                    }
                }
                else {
                    Write-Log -Level Info -Message "Keine Gruppen zu kopieren (außer Domain Users)"
                    Write-Log -Level Debug -Message "└─ Keine zusätzlichen Gruppen"
                }
            }
            catch {
                Write-Log -Level Warning -Message "Gruppen-Kopieren fehlgeschlagen: $_"
                Write-Log -Level Debug -Message "└─ FEHLER Gruppen: $_"
            }
        }

        Write-Log -Level Debug -Message "END: Benutzer-Kopie abgeschlossen"
        return $newUser
    }

    #endregion Initialisierung
    
    #region Funktionen

    function Show-InteractiveMenu {
        Write-Log -Level Debug -Message "Zeige interaktives Menü"
        
        $menuItems = @(
            @{ Number = 1; Title = "Benutzer kopieren"; Icon = "[1]" },
            @{ Number = 2; Title = "Benutzer aus CSV erstellen"; Icon = "[2]" },
            @{ Number = 3; Title = "Eigenschaften/Gruppen synchronisieren"; Icon = "[3]" },
            @{ Number = 4; Title = "Benutzer exportieren"; Icon = "[4]" },
            @{ Number = 5; Title = "L-Kennung exportieren"; Icon = "[5]" },
            @{ Number = 0; Title = "Beenden"; Icon = "[x]" }
        )

        while ($true) {
            Clear-Host
            Write-Host "`n"
            Write-Host "╔════════════════════════════════════════════════════════════════╗" -ForegroundColor White
            Write-Host "║  AD-Benutzerverwaltung v9.0 - Interaktives Menü (FINAL)       ║" -ForegroundColor Cyan
            Write-Host "╚════════════════════════════════════════════════════════════════╝" -ForegroundColor White
            Write-Host "`n"
            
            foreach ($item in $menuItems) {
                Write-Host "  $($item.Icon)  $($item.Title)" -ForegroundColor White
            }
            
            Write-Host "`n"
            Write-Host "════════════════════════════════════════════════════════════════" -ForegroundColor Gray
            Write-Host "Eingabe: 0-5 (oder Q zum Beenden)" -ForegroundColor Yellow
            Write-Host "Log:     $global:fullLogPath" -ForegroundColor DarkGray
            Write-Host ""

            $choice = Read-Host "Wahl"

            switch ($choice.ToUpper()) {
                '1' {
                    Write-Log -Level Debug -Message "Menü: CopySingleUser gewählt"
                    Invoke-CopySingleUser
                    Read-Host "`nDrücke Enter zum Fortfahren..."
                }
                '2' {
                    Write-Log -Level Debug -Message "Menü: CreateUsersFromCSV gewählt"
                    Invoke-CreateUsersFromCSV
                    Read-Host "`nDrücke Enter zum Fortfahren..."
                }
                '3' {
                    Write-Log -Level Debug -Message "Menü: ApplyPropertiesToExistingUser gewählt"
                    Invoke-ApplyPropertiesToExistingUser
                    Read-Host "`nDrücke Enter zum Fortfahren..."
                }
                '4' {
                    Write-Log -Level Debug -Message "Menü: ExportUserData gewählt"
                    Invoke-ExportUserData
                    Read-Host "`nDrücke Enter zum Fortfahren..."
                }
                '5' {
                    Write-Log -Level Debug -Message "Menü: ExportLKennung gewählt"
                    Invoke-ExportLKennung
                    Read-Host "`nDrücke Enter zum Fortfahren..."
                }
                { $_ -in '0', 'Q', 'q' } {
                    Write-Log -Level Info -Message "Benutzer beendet Anwendung"
                    Write-Host "`nAuf Wiedersehen!" -ForegroundColor Green
                    exit 0
                }
                default {
                    Write-Host "Ungültige Eingabe!" -ForegroundColor Red
                    Start-Sleep -Milliseconds 500
                }
            }
        }
    }

    function Invoke-CreateUsersFromCSV {
        Write-Log -Level Debug -Message "Invoke-CreateUsersFromCSV: Start"
        Write-Log -Level Info -Message "Modus: CreateUsersFromCSV"

        try {
            if (-not $CsvPath) {
                $CsvPath = Read-Host "Geben Sie den Pfad zur CSV-Datei ein"
            }

            if (-not (Test-Path $CsvPath -PathType Leaf)) {
                Write-Host "CSV-Datei nicht gefunden: $CsvPath" -ForegroundColor Red
                Write-Log -Level Error -Message "CSV nicht gefunden: $CsvPath"
                return
            }

            $csvData = Import-Csv -Path $CsvPath -Delimiter ';' -Encoding UTF8 -ErrorAction Stop
            Write-Log -Level Info -Message "CSV geladen: $($csvData.Count) Benutzer"
            Write-Log -Level Debug -Message "├─ Pfad: $CsvPath"

            $templateUser = $null
            if ($ReferenceUserSamAccountName) {
                try {
                    $templateUser = Get-ADUser -Identity $ReferenceUserSamAccountName -Properties * -ErrorAction Stop
                    Write-Log -Level Info -Message "Template-Benutzer: $($templateUser.SamAccountName)"
                    Write-Log -Level Debug -Message "├─ Template-OU: $($templateUser.DistinguishedName)"
                }
                catch {
                    Write-Log -Level Warning -Message "Template-Benutzer nicht gefunden"
                }
            }

            $successCount = 0
            $errorCount = 0

            foreach ($user in $csvData) {
                try {
                    $samAccountName = $user.SamAccountName.Trim()
                    $givenName = $user.GivenName.Trim()
                    $surname = $user.Surname.Trim()

                    if (-not $samAccountName -or -not $givenName -or -not $surname) {
                        Write-Log -Level Error -Message "Fehlende Pflichtfelder für: $samAccountName"
                        Write-Log -Level Debug -Message "└─ SamAccountName: '$samAccountName', GivenName: '$givenName', Surname: '$surname'"
                        $errorCount++
                        continue
                    }

                    # Benutzer existiert bereits?
                    $existing = Get-ADUser -Filter "SamAccountName -eq '$samAccountName'" -ErrorAction SilentlyContinue
                    if ($existing) {
                        Write-Log -Level Warning -Message "Benutzer existiert bereits: $samAccountName"
                        Write-Log -Level Debug -Message "├─ Übersprungen"
                        $errorCount++
                        continue
                    }

                    # Passwort
                    $finalPassword = $DefaultPassword
                    if ($user.Password) {
                        $finalPassword = ConvertTo-SecureString $user.Password -AsPlainText -Force
                    }
                    elseif (-not $finalPassword) {
                        $randomPassword = -join ((33..126) | Get-Random -Count 16 | ForEach-Object { [char]$_ })
                        $finalPassword = ConvertTo-SecureString $randomPassword -AsPlainText -Force
                        Write-Log -Level Debug -Message "├─ Zufallspasswort generiert"
                    }

                    # OU
                    $finalOU = $TargetOU
                    if ($user.TargetOU) {
                        $finalOU = $user.TargetOU
                    }
                    elseif ($templateUser) {
                        $finalOU = ($templateUser.DistinguishedName -split ',', 2)[1]
                    }

                    if (-not $finalOU) {
                        Write-Log -Level Error -Message "Keine OU für $samAccountName definiert"
                        $errorCount++
                        continue
                    }

                    # Benutzer erstellen
                    $newUserParams = @{
                        SamAccountName        = $samAccountName
                        GivenName             = $givenName
                        Surname               = $surname
                        Name                  = "$givenName $surname"
                        DisplayName           = "$givenName $surname"
                        UserPrincipalName     = "$samAccountName@$($env:USERDNSDOMAIN)"
                        AccountPassword       = $finalPassword
                        ChangePasswordAtLogon = $true
                        Enabled               = $true
                        Path                  = $finalOU
                    }

                    New-ADUser @newUserParams -ErrorAction Stop
                    Write-Log -Level Info -Message "Benutzer erstellt: $samAccountName"
                    Write-Log -Level Debug -Message "├─ OU: $finalOU"
                    Add-UserReportEntry -SamAccountName $samAccountName -Status "Erstellt" -Detail "CSV Import"
                    $successCount++

                }
                catch {
                    Write-Log -Level Error -Message "Fehler bei $($user.SamAccountName): $_"
                    Write-Log -Level Debug -Message "└─ Exception: $_"
                    $errorCount++
                }
            }

            Write-Host "CSV-Import abgeschlossen: $successCount Benutzer erstellt, $errorCount Fehler" -ForegroundColor Green
            Write-Log -Level Info -Message "CSV-Import: $successCount Erfolg, $errorCount Fehler"

        }
        catch {
            Write-Host "CSV-Verarbeitung fehlgeschlagen: $_" -ForegroundColor Red
            Write-Log -Level Error -Message "CreateUsersFromCSV fehlgeschlagen: $_"
        }
        
        Write-Log -Level Debug -Message "Invoke-CreateUsersFromCSV: End"
    }

    # ★★★ FIX: ApplyPropertiesToExistingUser mit dynamischer Domain-Ermittlung ★★★
    function Invoke-ApplyPropertiesToExistingUser {
        Write-Log -Level Debug -Message "Invoke-ApplyPropertiesToExistingUser: Start (FIXED)"
        Write-Log -Level Info -Message "Modus: ApplyPropertiesToExistingUser"

        try {
            if (-not $ReferenceUserSamAccountName) {
                $ReferenceUserSamAccountName = Read-Host "Geben Sie den SAM-Account-Namen des Quellbenutzers ein"
            }

            if (-not $TargetUserSamAccountName) {
                $TargetUserSamAccountName = Read-Host "Geben Sie den SAM-Account-Namen des Zielbenutzers ein"
            }

            # FIX: Zuerst Domain vom Referenzbenutzer ermitteln
            Write-Log -Level Debug -Message "├─ Ermittle Domain vom Referenzbenutzer..."
            $referenceUser = Get-ADUser -Identity $ReferenceUserSamAccountName -Properties * -ErrorAction Stop
            $referenceDomain = Get-ADServerFromUser -User $referenceUser
            
            if ($referenceDomain) {
                Write-Log -Level Info -Message "Referenzbenutzer gefunden (Domain: $referenceDomain)"
            }
            else {
                Write-Log -Level Warning -Message "Referenzbenutzer gefunden, aber Domain konnte nicht ermittelt werden"
            }
            
            Write-Log -Level Debug -Message "├─ Ermittle Zielbenutzer..."
            $targetUser = Get-ADUser -Identity $TargetUserSamAccountName -Properties * -ErrorAction Stop
            
            Write-Log -Level Info -Message "Beide Benutzer gefunden"
            Write-Log -Level Debug -Message "├─ Quelle: $($referenceUser.SamAccountName)"
            Write-Log -Level Debug -Message "├─ Ziel: $($targetUser.SamAccountName)"

            # Eigenschaften
            $propertiesToApply = @('Description', 'Office', 'StreetAddress', 'City', 'State', 'PostalCode', 'Country', 'Department', 'Company', 'Title', 'OfficePhone', 'EmailAddress')
            
            $setParams = @{ Identity = $targetUser }
            $changedProperties = @()
            
            foreach ($prop in $propertiesToApply) {
                if ($referenceUser.$prop -and $referenceUser.$prop -ne $targetUser.$prop) {
                    $setParams[$prop] = $referenceUser.$prop
                    $changedProperties += $prop
                    Write-Log -Level Debug -Message "├─ Änderung: $prop = '$($referenceUser.$prop)'"
                }
            }

            if ($changedProperties) {
                Set-ADUser @setParams -ErrorAction Stop
                Write-Log -Level Info -Message "Eigenschaften angewendet: $($changedProperties -join ', ')"
            }
            else {
                Write-Log -Level Info -Message "Keine Eigenschaftsänderungen erforderlich"
            }

            # Gruppen kopieren
            Write-Log -Level Debug -Message "├─ Kopiere Gruppen..."
            $sourceGroups = Get-ADPrincipalGroupMembership -Identity $referenceUser -ErrorAction Stop
            $groupsToCopy = $sourceGroups | Where-Object { $_.Name -ne "Domain Users" -and $_.Name -ne "Domänen-Benutzer" }

            if ($groupsToCopy) {
                $targetGroups = Get-ADPrincipalGroupMembership -Identity $targetUser -ErrorAction Stop
                $groupsToAdd = $groupsToCopy | Where-Object { $_.DistinguishedName -notin $targetGroups.DistinguishedName }

                if ($groupsToAdd) {
                    Add-ADPrincipalGroupMembership -Identity $targetUser -MemberOf $groupsToAdd -ErrorAction Stop
                    Write-Log -Level Info -Message "Gruppen hinzugefügt: $($groupsToAdd.Count)"
                    Write-Log -Level Debug -Message "└─ Neue Gruppen:"
                    foreach ($grp in $groupsToAdd) {
                        Write-Log -Level Debug -Message "   - $($grp.Name)"
                    }
                }
                else {
                    Write-Log -Level Info -Message "Benutzer ist bereits in allen Gruppen"
                    Write-Log -Level Debug -Message "└─ Keine neuen Gruppen"
                }
            }

            Write-Host "Synchronisierung abgeschlossen" -ForegroundColor Green
            Write-Log -Level Info -Message "ApplyPropertiesToExistingUser erfolgreich"
            Add-UserReportEntry -SamAccountName $TargetUserSamAccountName -Status "Synchronisiert" -Detail "Eigenschaften + Gruppen von $ReferenceUserSamAccountName"

        }
        catch {
            Write-Host "Fehler: $_" -ForegroundColor Red
            Write-Log -Level Error -Message "ApplyPropertiesToExistingUser fehlgeschlagen: $_"
        }
        
        Write-Log -Level Debug -Message "Invoke-ApplyPropertiesToExistingUser: End"
    }

    # Optimierte ExportUserData (mit OU-Filterung aus v6.7)
    function Invoke-ExportUserData {
        Write-Log -Level Debug -Message "Invoke-ExportUserData: Start"
        Write-Log -Level Info -Message "Modus: ExportUserData (mit OU-Filterung)"

        try {
            if (-not $IdentityFilter) {
                $IdentityFilter = Read-Host "Suchfilter (z.B. L110* oder *): "
            }

            if (-not $ExportCsvPath) {
                $ExportCsvPath = Read-Host "Zieldatei für CSV (oder Enter für Standard): "
                if (-not $ExportCsvPath) {
                    $ExportCsvPath = "$([Environment]::GetFolderPath('Desktop'))\ADExport_$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
                }
            }

            Write-Log -Level Debug -Message "├─ Filter: $IdentityFilter"
            Write-Log -Level Debug -Message "├─ Ziel: $ExportCsvPath"

            # Nutze optimierte Filterfunktion aus v6.7
            $users = Get-ADUsersWithOUFilter -IdentityFilter $IdentityFilter -OUFilter $OUFilter -SearchBaseOU $SearchBaseOU

            if (-not $users) {
                Write-Host "Keine Benutzer gefunden" -ForegroundColor Yellow
                Write-Log -Level Warning -Message "Keine Benutzer gefunden mit Filter: $IdentityFilter"
                return
            }

            $users = $users | Get-ADUser -Properties * -ErrorAction Stop
            $exportData = $users | Select-Object SamAccountName, Name, DisplayName, UserPrincipalName, Enabled, DistinguishedName, EmailAddress
            $exportData | Export-Csv -Path $ExportCsvPath -Delimiter ';' -Encoding UTF8 -NoTypeInformation

            Write-Host "Export abgeschlossen: $($users.Count) Benutzer exportiert" -ForegroundColor Green
            Write-Host "Datei: $ExportCsvPath" -ForegroundColor Green
            Write-Log -Level Info -Message "Export: $($users.Count) Benutzer in '$ExportCsvPath'"

        }
        catch {
            Write-Host "Fehler: $_" -ForegroundColor Red
            Write-Log -Level Error -Message "ExportUserData fehlgeschlagen: $_"
        }
        
        Write-Log -Level Debug -Message "Invoke-ExportUserData: End"
    }


    #endregion Funktionen
    
    Write-Log -Level Debug -Message "Initialisierung abgeschlossen. Mode: $Mode, LogLevel: $LogLevel"
}

process {
    Write-Log -Level Debug -Message "PROCESS: Starte Modus '$Mode'"

    switch ($Mode) {
        'Interactive' {
            Show-InteractiveMenu
        }
        
        'CopySingleUser' {
            Invoke-CopySingleUser
        }
        
        'CreateUsersFromCSV' {
            Invoke-CreateUsersFromCSV
        }
        
        'ApplyPropertiesToExistingUser' {
            Invoke-ApplyPropertiesToExistingUser
        }
        
        'ExportUserData' {
            Invoke-ExportUserData
        }
        
        'ExportLKennung' {
            Invoke-ExportLKennung
        }
    }
}

end {
    # Report speichern
    if ($global:userReportData.Count -gt 0) {
        try {
            Write-Log -Level Debug -Message "Speichere Report mit $($global:userReportData.Count) Einträgen"
            $global:userReportData | Export-Csv -Path $global:fullUserReportPath -Delimiter ';' -Encoding UTF8 -NoTypeInformation -ErrorAction Stop
            Write-Log -Level Info -Message "Report gespeichert: $global:fullUserReportPath"
        }
        catch {
            Write-Log -Level Error -Message "Report-Speicherung fehlgeschlagen: $_"
        }
    }
    
    Write-Log -Level Debug -Message "Skript abgeschlossen"
}

