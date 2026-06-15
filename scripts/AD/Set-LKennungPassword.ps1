# FILE: scripts\AD\Set-LKennungPassword.ps1
# PURPOSE: Set passwords and update AD user properties with UseNetUser and Silent parameters (V4)
# DEPENDS ON: ActiveDirectory module
# DEPENDED ON BY: None
# LAST MODIFIED: 2026-06-15
# MODIFIED BY: Systems Administration
# CHANGE SUMMARY: Update Set-LKennungPassword.ps1 to support net user bypass and silent preference (V4)
# BRANCH: main

#Requires -Version 5.1
#Requires -Modules ActiveDirectory

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter()]
    [string[]]$OUNames = @('81', '82'), 

    [Parameter()]
    [string]$ReportPath = "C:\Daten\Benutzerbericht.csv",

    [Parameter()]
    [switch]$Force,

    [Parameter(HelpMessage = 'Verwendet "net user" für den Passwort-Reset. Andernfalls wird Set-ADUser genutzt.')]
    [switch]$UseNetUser,

    [Parameter(HelpMessage = 'Unterdrückt die ausführliche Konsolenausgabe (Verbose).')]
    [switch]$Silent
)

begin {
    #region Initialisierung und Lokalisierung
    if ($Silent) {
        $VerbosePreference = 'SilentlyContinue'
    } else {
        $VerbosePreference = 'Continue'
    }
    Write-Verbose "Skriptstart am $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')."
    Write-Verbose "PowerShell-Version: $($PSVersionTable.PSVersion)."
    Write-Verbose "Verwendete Parameter: OUNames='$($OUNames -join ', ')', ReportPath='$ReportPath', Force=$Force, UseNetUser=$UseNetUser, Silent=$Silent."

    $ErrorActionPreference = 'Stop'

    # Deutsche Lokalisierung einstellen für Datums-/Zeitformate und Zahlen
    Write-Verbose "Stelle Kultur auf 'de-DE' ein."
    $culture = [System.Globalization.CultureInfo]::GetCultureInfo('de-DE')
    [System.Threading.Thread]::CurrentThread.CurrentCulture = $culture
    [System.Threading.Thread]::CurrentThread.CurrentUICulture = $culture

    $global:OperationResults = [System.Collections.Generic.List[object]]::new()
    $PasswordSetDate = Get-Date

    #region Hilfsfunktionen
    function Write-Header {
        param([string]$Title)
        if (-not $Silent) {
            Write-Host "`n$(('=' * 80))" -ForegroundColor Cyan
            Write-Host " $Title " -ForegroundColor Cyan
            Write-Host "$(('=' * 80))`n" -ForegroundColor Cyan
        }
        Write-Verbose "Header angezeigt: '$Title'."
    }
    #endregion Hilfsfunktionen

    # Dynamische Ermittlung des Domänen-DN
    Write-Verbose "Ermittle den DistinguishedName der aktuellen AD-Domäne."
    try {
        $domain = (Get-ADDomain -ErrorAction Stop).DistinguishedName
        Write-Verbose "AD-Domäne erfolgreich ermittelt: '$domain'."
    } catch {
        Write-Host "Fehler beim Domänenzugriff: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }

    if (-not ($PSBoundParameters.ContainsKey('ReportPath'))) {
        Write-Header -Title "INTERAKTIVE EINGABE: BERICHTSPFAD"
        $inputReportPath = Read-Host "Bitte geben Sie den vollständigen Pfad für den Berichts-Export an. Leer lassen für Standard: '$ReportPath'"
        if (-not [string]::IsNullOrWhiteSpace($inputReportPath)) {
            $ReportPath = $inputReportPath.Trim()
        }
    }

    try {
        $reportDirectory = Split-Path $ReportPath -Parent
        if (-not (Test-Path $reportDirectory)) {
            New-Item -ItemType Directory -Path $reportDirectory -Force | Out-Null
        }
    } catch {
        Write-Host "Fehler beim Überprüfen/Erstellen des Berichtspfad-Verzeichnisses '$reportDirectory': $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
    #endregion Initialisierung und Lokalisierung
}

process {
    try {
        # 1. OU-SUCHE ---------------------------------------------------------
        Write-Header -Title "SCHRITT 1: ORGANISATIONSEINHEITEN-SUCHE"
        Write-Host "Suche nach OUs: $($OUNames -join ', ')" -ForegroundColor Yellow

        $targetOUs = @()
        foreach ($name in $OUNames) {
            try {
                $ou = Get-ADOrganizationalUnit -Filter "Name -eq '$name'" `
                    -SearchBase $domain `
                    -SearchScope Subtree `
                    -ErrorAction Stop
                $targetOUs += $ou
            } catch {
                Write-Warning "OU '$name' konnte nicht gefunden werden: $($_.Exception.Message)."
            }
        }

        if (-not $targetOUs.Count) {
            Write-Host "Keine der OUs gefunden. Skript wird beendet." -ForegroundColor Red
            return
        }

        # 2. BENUTZERSUCHE ----------------------------------------------------
        Write-Header -Title "SCHRITT 2: BENUTZERSUCHE"
        $users = @()
        foreach ($ou in $targetOUs) {
            try {
                $foundUsers = Get-ADUser -LDAPFilter "(|(sAMAccountName=L110*)(sAMAccountName=L114*))" `
                    -SearchBase $ou.DistinguishedName `
                    -Properties Name, SamAccountName, Enabled, LastLogonDate, DistinguishedName, `
                                PasswordNeverExpires, CannotChangePassword, PasswordLastSet, `
                                LastLogonTimestamp, UserAccountControl, PasswordChangeRequired `
                    -SearchScope Subtree `
                    -ErrorAction Stop
                $users += $foundUsers
            } catch {
                Write-Warning "Fehler beim Suchen von Benutzern in OU '$($ou.Name)': $($_.Exception.Message)"
            }
        }

        if (-not $users.Count) {
            Write-Host "Keine passenden Benutzer gefunden." -ForegroundColor Yellow
            return
        }

        # 3. BENUTZERANZEIGE --------------------------------------------------
        Write-Header -Title "GEFUNDENE BENUTZER"
        $users | Select-Object @{l='Benutzername';e={$_.Name}},
                                @{l='Anmeldename';e={$_.SamAccountName}},
                                @{l='Aktiviert';e={if($_.Enabled){'Ja'}else{'Nein'}}},
                                @{l='Passwort läuft nie ab';e={if($_.PasswordNeverExpires){'Ja'}else{'Nein'}}},
                                @{l='Passwort bei nächster Anmeldung ändern';e={if($_.PasswordChangeRequired){'Ja'}else{'Nein'}}},
                                @{l='Letzte Anmeldung';e={$_.LastLogonDate.ToString('dd.MM.yyyy HH:mm:ss', $culture)}},
                                @{l='OU-Pfad';e={$_.DistinguishedName}} | Format-Table -AutoSize

        # 4. PASSWORTHINWEIS --------------------------------------------------
        Write-Header -Title "PASSWORTINFORMATION"
        Write-Host "Geplantes Passwort-Änderungsdatum für Protokollierung: $($PasswordSetDate.ToString('dd.MM.yyyy HH:mm:ss'))" -ForegroundColor Magenta

        # 5. BESTÄTIGUNG ------------------------------------------------------
        Write-Header -Title "BESTÄTIGUNG"
        if (-not $Force) {
            $confirmation = Read-Host "Möchten Sie die Änderungen für die gefundenen Benutzer durchführen? (J/N)"
            if ($confirmation -notin @('J','j')) {
                Write-Host "Operation vom Benutzer abgebrochen." -ForegroundColor Yellow
                return
            }
        }

        # 6. PASSWORTVERWALTUNG -----------------------------------------------
        Write-Header -Title "PASSWORTZURÜCKSETZUNG"
        $securePass = Read-Host "Neues Passwort eingeben" -AsSecureString

        $plainPass = $null
        if ($UseNetUser) {
            $plainPass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
                [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePass)
            )
        }

        # 7. BENUTZERVERARBEITUNG ---------------------------------------------
        foreach ($user in $users) {
            $resultTemplate = [PSCustomObject]@{
                Benutzername            = $user.Name
                Anmeldename             = $user.SamAccountName
                Aktiviert               = $user.Enabled
                LetzteAnmeldung         = $user.LastLogonDate.ToString('dd.MM.yyyy HH:mm:ss', $culture)
                OU_Pfad                 = $user.DistinguishedName
                PasswortNieAblauf       = $user.PasswordNeverExpires
                PasswortAendernNoetig   = $user.PasswordChangeRequired
                OperationStatus         = "Nicht durchgeführt"
                FehlerCode              = $null
                FehlerMeldung           = $null
                PasswortDatum           = $null
            }

            try {
                if ($PSCmdlet.ShouldProcess($user.SamAccountName, "Passwort zurücksetzen und Attribute aktualisieren")) {
                    Write-Host "`nVerarbeite Benutzer: $($user.SamAccountName)" -ForegroundColor Cyan

                    if ($UseNetUser) {
                        $output = & net user $user.SamAccountName $plainPass /DOMAIN /ACTIVE:YES 2>&1
                        if ($LASTEXITCODE -ne 0) {
                            throw "NetUser-Fehler ($LASTEXITCODE): $output"
                        }

                        Set-ADUser -Identity $user.DistinguishedName `
                                   -PasswordNeverExpires $true `
                                   -CannotChangePassword $true `
                                   -PasswordChangeRequired $false `
                                   -Replace @{lastLogonTimestamp = (Get-Date).ToFileTime()} `
                                   -ErrorAction Stop
                    } else {
                        Set-ADUser -Identity $user.DistinguishedName `
                                   -NewPassword $securePass `
                                   -PasswordNeverExpires $true `
                                   -CannotChangePassword $true `
                                   -PasswordChangeRequired $false `
                                   -Replace @{lastLogonTimestamp = (Get-Date).ToFileTime()} `
                                   -ErrorAction Stop
                    }

                    $resultTemplate.OperationStatus = "Erfolgreich"
                    $resultTemplate.PasswortDatum = (Get-Date).ToString('dd.MM.yyyy HH:mm:ss', $culture)
                    Write-Host "Benutzer '$($user.SamAccountName)' erfolgreich aktualisiert." -ForegroundColor Green
                } else {
                    $resultTemplate.OperationStatus = "Übersprungen (ShouldProcess)"
                    Write-Host "Verarbeitung von Benutzer '$($user.SamAccountName)' übersprungen." -ForegroundColor Yellow
                }
            }
            catch {
                $resultTemplate.OperationStatus = "Fehlgeschlagen"
                $resultTemplate.FehlerCode = $_.Exception.HResult
                $resultTemplate.FehlerMeldung = $_.Exception.Message
                $resultTemplate.PasswortDatum = $null
                Write-Host "Fehler bei der Verarbeitung von Benutzer '$($user.SamAccountName)': $($_.Exception.Message)" -ForegroundColor Red
            }
            finally {
                $global:OperationResults.Add($resultTemplate)
            }
        }
    }
    catch {
        Write-Host "KRITISCHER FEHLER WÄHREND DER AUSFÜHRUNG: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
    finally {
        # 8. CSV-EXPORT -------------------------------------------------------
        Write-Header -Title "CSV-EXPORT"
        if ($global:OperationResults.Count -gt 0) {
            try {
                $global:OperationResults | Export-Csv -Path $ReportPath -Delimiter ";" -Encoding UTF8 -NoTypeInformation -ErrorAction Stop
                Write-Host "Bericht erfolgreich erstellt: $ReportPath" -ForegroundColor Green
            } catch {
                Write-Host "Fehler beim Exportieren des Berichts nach '$ReportPath': $($_.Exception.Message)" -ForegroundColor Red
            }
        }

        if ($plainPass) {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR(
                [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePass)
            )
            Remove-Variable plainPass -ErrorAction SilentlyContinue
        }
        Remove-Variable securePass -ErrorAction SilentlyContinue
    }
}

end {
    # 9. ZUSAMMENFASSUNG ------------------------------------------------------
    Write-Header -Title "ZUSAMMENFASSUNG"
    $successCount = ($global:OperationResults | Where-Object { $_.OperationStatus -eq "Erfolgreich" }).Count
    $skippedCount = ($global:OperationResults | Where-Object { $_.OperationStatus -eq "Übersprungen (ShouldProcess)" }).Count
    $errorCount = ($global:OperationResults | Where-Object { $_.OperationStatus -eq "Fehlgeschlagen" }).Count

    Write-Host @"
    Zusammenfassung der Operationen:
    Verarbeitete Benutzer: $($global:OperationResults.Count)
    Erfolgreich durchgeführt: $successCount
    Übersprungen:             $skippedCount
    Fehlgeschlagen:           $errorCount
"@ -ForegroundColor DarkCyan

    if ($errorCount -gt 0) {
        exit 2
    } else {
        exit 0
    }
}
