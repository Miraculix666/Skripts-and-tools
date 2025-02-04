<#
.SYNOPSIS
    Erstellt AD-Benutzer als exakte Kopien von Vorlagenbenutzern mit automatischer Gruppenübernahme.

.DESCRIPTION
    Dieses Skript ermöglicht die Massenerstellung von AD-Benutzern basierend auf vorhandenen Vorlagen.
    Es kopiert:
    - Alle Standardattribute (außer systemeigene Eigenschaften)
    - Gruppenmitgliedschaften
    - Organisational Unit (OU)
    - Benutzerkonto-Einstellungen

    Unterstützte Modi:
    - Interaktive Eingabe
    - Parameterbasierte Einzelerstellung
    - CSV-Batch-Verarbeitung

.PARAMETER CsvPath
    Pfad zur CSV-Datei mit Batch-Erstellungsdaten (Format siehe NOTES)

.PARAMETER TemplateUser
    SAMAccountName des Vorlagenbenutzers

.PARAMETER NewUserName
    SAMAccountName für den neuen Benutzer

.PARAMETER NewUserPassword
    Initialpasswort für den neuen Benutzer

.PARAMETER Verify
    Aktiviert die Post-Erstellungsüberprüfung

.PARAMETER WhatIf
    Simuliert die Aktionen ohne Änderungen am AD

.EXAMPLE
    # Interaktiver Modus
    .\New-ADUserClone.ps1

    # CSV-Batch-Modus mit Überprüfung
    .\New-ADUserClone.ps1 -CsvPath .\users.csv -Verify

    # Einzelerstellung mit Parameter
    .\New-ADUserClone.ps1 -TemplateUser "jdoe" -NewUserName "asmith" -NewUserPassword "P@ssw0rd123!"

.NOTES
    Erforderliche CSV-Struktur:
    TemplateUser,NewUserName,NewUserPassword,Description,Department,Title
    jdoe,asmith,P@ssw0rd123!,Developer,IT,Junior Developer
    msmith,bjones,SecurePwd456!,Manager,Sales,Sales Lead

    Version: 2.1
    Autor: IT-Abteilung
    Letzte Änderung: 2024-03-15
#>

[CmdletBinding(DefaultParameterSetName = 'Interactive', SupportsShouldProcess = $true)]
param (
    [Parameter(ParameterSetName = 'CSV', Mandatory = $true)]
    [ValidateScript({
        if (-not (Test-Path $_)) { throw "CSV-Datei nicht gefunden" }
        $true
    })]
    [string]$CsvPath,

    [Parameter(ParameterSetName = 'Single', Mandatory = $true)]
    [string]$TemplateUser,

    [Parameter(ParameterSetName = 'Single', Mandatory = $true)]
    [string]$NewUserName,

    [Parameter(ParameterSetName = 'Single')]
    [string]$NewUserPassword,

    [Parameter()]
    [switch]$Verify,

    [Parameter()]
    [switch]$WhatIf
)

begin {
    #region Modulimport und Initialisierung
    Import-Module ActiveDirectory -ErrorAction Stop

    # Logging-Konfiguration
    $logFile = Join-Path $PSScriptRoot "ADUserClone_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    
    function Write-Log {
        param(
            [string]$Message,
            [ValidateSet('INFO','WARNING','ERROR','SUCCESS')]
            [string]$Level = 'INFO'
        )
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "[$timestamp][$Level] $Message"
        Add-Content -Path $logFile -Value $logEntry
        Write-Host ("{0}: {1}" -f $Level.PadRight(7), $Message) -ForegroundColor (
            switch ($Level) {
                'INFO'    { 'Gray' }
                'WARNING' { 'Yellow' }
                'ERROR'   { 'Red' }
                'SUCCESS' { 'Green' }
            }
        )
    }
    #endregion

    #region Hilfsfunktionen
    function Get-ADUserSafe {
        param([string]$SamAccountName)
        try {
            Get-ADUser -Identity $SamAccountName -Properties * -ErrorAction Stop
        }
        catch {
            Write-Log "Benutzer $SamAccountName nicht gefunden: $_" -Level ERROR
            return $null
        }
    }

    function Copy-ADGroupMembership {
        param(
            [string]$SourceUser,
            [string]$TargetUser
        )
        try {
            $groups = Get-ADUser -Identity $SourceUser -Properties MemberOf | 
                     Select-Object -ExpandProperty MemberOf
            $count = 0
            
            foreach ($group in $groups) {
                if ($PSCmdlet.ShouldProcess($group, "Add group membership")) {
                    Add-ADGroupMember -Identity $group -Members $TargetUser -ErrorAction Stop
                    $count++
                }
            }
            Write-Log "$count Gruppenmitgliedschaften kopiert" -Level SUCCESS
        }
        catch {
            Write-Log "Fehler beim Gruppenkopieren: $_" -Level ERROR
            throw
        }
    }
    #endregion
}

process {
    #region Hauptverarbeitung
    try {
        $createdUsers = @()

        # CSV-Batch-Verarbeitung
        if ($PSCmdlet.ParameterSetName -eq 'CSV') {
            Write-Log "Starte CSV-Modus mit Datei: $CsvPath" -Level INFO
            $users = Import-Csv -Path $CsvPath -Encoding UTF8

            # CSV-Validierung
            $requiredColumns = @('TemplateUser', 'NewUserName', 'NewUserPassword')
            $missingColumns = $requiredColumns | Where-Object { $_ -notin $users[0].PSObject.Properties.Name }
            if ($missingColumns) {
                throw "Fehlende Spalten in CSV: $($missingColumns -join ', ')"
            }

            foreach ($user in $users) {
                # Duplikatsprüfung
                if (Get-ADUserSafe -SamAccountName $user.NewUserName) {
                    Write-Log "Überspringe vorhandenen Benutzer: $($user.NewUserName)" -Level WARNING
                    continue
                }

                # Passwortvalidierung
                if ($user.NewUserPassword.Length -lt 12 -or 
                    -not ($user.NewUserPassword -match '\d') -or
                    -not ($user.NewUserPassword -match '[A-Z]') -or
                    -not ($user.NewUserPassword -match '[a-z]')) {
                    Write-Log "Passwort für $($user.NewUserName) entspricht nicht den Richtlinien" -Level ERROR
                    continue
                }

                # Benutzererstellung
                $template = Get-ADUserSafe -SamAccountName $user.TemplateUser
                if (-not $template) {
                    continue
                }

                $userParams = @{
                    Instance           = $template
                    SamAccountName     = $user.NewUserName
                    UserPrincipalName  = "$($user.NewUserName)@$((Get-ADDomain).DNSRoot)"
                    Name               = $user.NewUserName
                    AccountPassword    = ConvertTo-SecureString $user.NewUserPassword -AsPlainText -Force
                    Enabled            = $true
                    ChangePasswordAtLogon = $true
                }

                if ($PSCmdlet.ShouldProcess($user.NewUserName, "Create AD User")) {
                    $newUser = New-ADUser @userParams -PassThru
                    Copy-ADGroupMembership -SourceUser $user.TemplateUser -TargetUser $user.NewUserName
                    $createdUsers += $newUser
                    Write-Log "Benutzer $($user.NewUserName) erfolgreich erstellt" -Level SUCCESS
                }
            }
        }
        # Einzelbenutzer-Modus
        else {
            # Interaktive Abfrage fehlender Parameter
            if ([string]::IsNullOrEmpty($TemplateUser)) {
                $TemplateUser = Read-Host "Vorlagenbenutzer (SAMAccountName)"
            }
            
            if ([string]::IsNullOrEmpty($NewUserName)) {
                $NewUserName = Read-Host "Neuer Benutzername (SAMAccountName)"
            }

            if ([string]::IsNullOrEmpty($NewUserPassword)) {
                $securePass = Read-Host "Passwort" -AsSecureString
                $NewUserPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
                    [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePass)
                )
            }

            # Validierungen
            $template = Get-ADUserSafe -SamAccountName $TemplateUser
            if (-not $template) {
                throw "Vorlagenbenutzer nicht gefunden"
            }

            if (Get-ADUserSafe -SamAccountName $NewUserName) {
                throw "Benutzer $NewUserName existiert bereits"
            }

            # Benutzererstellung
            $userParams = @{
                Instance           = $template
                SamAccountName     = $NewUserName
                UserPrincipalName  = "$NewUserName@$((Get-ADDomain).DNSRoot)"
                Name               = $NewUserName
                AccountPassword    = ConvertTo-SecureString $NewUserPassword -AsPlainText -Force
                Enabled            = $true
                ChangePasswordAtLogon = $true
            }

            if ($PSCmdlet.ShouldProcess($NewUserName, "Create AD User")) {
                $newUser = New-ADUser @userParams -PassThru
                Copy-ADGroupMembership -SourceUser $TemplateUser -TargetUser $NewUserName
                $createdUsers += $newUser
                Write-Log "Benutzer $NewUserName erfolgreich erstellt" -Level SUCCESS
            }
        }
    }
    catch {
        Write-Log "KRITISCHER FEHLER: $_" -Level ERROR
        exit 1
    }
    #endregion
}

end {
    #region Post-Processing
    if ($Verify -and $createdUsers.Count -gt 0) {
        Write-Log "Starte Verifikation der erstellten Benutzer" -Level INFO
        foreach ($user in $createdUsers) {
            try {
                $adUser = Get-ADUser -Identity $user.SamAccountName -Properties *
                $template = Get-ADUser -Identity $userParams.Instance.SamAccountName -Properties *

                # Attributvergleich
                $compareProps = @(
                    'Enabled', 'AccountExpirationDate', 'PasswordNeverExpires',
                    'Department', 'Title', 'Office', 'Manager'
                )
                
                foreach ($prop in $compareProps) {
                    if ($adUser.$prop -ne $template.$prop) {
                        Write-Log "Abweichung bei $prop: Vorlage[$($template.$prop)] Neu[$($adUser.$prop)]" -Level WARNING
                    }
                }

                # Gruppenvergleich
                $templateGroups = Get-ADUser -Identity $template.SamAccountName -Properties MemberOf | 
                                 Select-Object -ExpandProperty MemberOf
                $newGroups = Get-ADUser -Identity $adUser.SamAccountName -Properties MemberOf | 
                            Select-Object -ExpandProperty MemberOf

                $diff = Compare-Object -ReferenceObject $templateGroups -DifferenceObject $newGroups
                if ($diff) {
                    Write-Log "Gruppenabweichungen gefunden: $($diff | Out-String)" -Level WARNING
                }
            }
            catch {
                Write-Log "Verifikationsfehler bei $($user.SamAccountName): $_" -Level ERROR
            }
        }
    }

    Write-Log "Prozess abgeschlossen. Log-Datei: $logFile" -Level INFO
    Invoke-Item $logFile
    #endregion
}
