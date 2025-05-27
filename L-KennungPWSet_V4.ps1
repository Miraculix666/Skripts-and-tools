<#
.SYNOPSIS
AD-Benutzerverwaltung mit erweitertem Logging und deutscher Lokalisierung.

.DESCRIPTION
Dieses Skript automatisiert die Verwaltung von Active Directory-Benutzerkonten.
Es sucht Benutzer in spezifischen Organisationseinheiten (OUs) basierend
auf Namensmustern, zeigt diese an, setzt Passwörter zurück (wahlweise via
'net user' oder 'Set-ADUser') und aktualisiert verschiedene Kontoeigenschaften,
einschließlich der Sicherstellung, dass das Flag 'Passwort bei nächster Anmeldung ändern'
nicht gesetzt ist. Alle Operationen werden detailliert protokolliert und in
einer CSV-Datei ausgegeben. Das Skript unterstützt deutsche Lokalisierung
für Datums-/Zeitformate und CSV-Trennzeichen.

.PARAMETER OUNames
Ein Array von OU-Namen, in denen nach Benutzern gesucht werden soll.
Standardmäßig auf '81' und '82' festgelegt.

.PARAMETER ReportPath
Der vollständige Pfad zur CSV-Berichtsdatei. Standardmäßig
'C:\Daten\Benutzerbericht.csv'.

.PARAMETER Force
Ein Switch-Parameter, der die Bestätigungsabfrage vor der Durchführung
von Änderungen unterdrückt.

.PARAMETER UseNetUser
Ein Switch-Parameter, der die Verwendung des 'net user'-Befehls für
den Passwort-Reset erzwingt. Standardmäßig wird 'Set-ADUser' verwendet.

.PARAMETER Silent
Ein Switch-Parameter, der die ausführliche Konsolenausgabe (Verbose) unterdrückt.
Standardmäßig ist die ausführliche Ausgabe aktiviert.

.EXAMPLE
# Standardausführung im interaktiven Modus
.\Invoke-ADUserMgmt.ps1

.EXAMPLE
# Ausführung mit spezifischer Berichtsdatei und ohne Bestätigung
.\Invoke-ADUserMgmt.ps1 -ReportPath "D:\Logs\ADBericht.csv" -Force

.EXAMPLE
# Ausführung mit 'net user' für den Passwort-Reset und stiller Ausgabe
.\Invoke-ADUserMgmt.ps1 -UseNetUser -Silent

.EXAMPLE
# Gemischter Modus: OUs werden interaktiv abgefragt, Berichtspfad ist vorgegeben
.\Invoke-ADUserMgmt.ps1 -ReportPath "C:\Temp\MeinBericht.csv"

.NOTES
Autor: PS-Coding
Version: 3.0.20250527
Erstellungsdatum: 27.05.2025
Letzte Änderung: 27.05.2025
Kompatibilität: PowerShell 5.1 (On-Premise Windows, Win10, Server 2019)

Referenzen:
- AI-discovered sources:
  - Microsoft Docs: Get-ADUser, Set-ADUser, Get-ADOrganizationalUnit, Get-ADDomain
  - Stack Overflow: SecureString handling, CultureInfo for localization
  - PowerShell Best Practices for Logging and Error Handling
- User-provided sources:
  - Original script provided by user
#>

#Requires -Version 5.1
#Requires -Modules ActiveDirectory

[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'Interactive')]
param(
    [Parameter(Mandatory = $false, HelpMessage = 'Geben Sie die Namen der OUs ein (z.B. "81", "82").')]
    [string[]]$OUNames = @('81', '82'),

    [Parameter(Mandatory = $false, HelpMessage = 'Geben Sie den vollständigen Pfad für den Berichts-Export an.')]
    [string]$ReportPath = "C:\Daten\Benutzerbericht.csv",

    [Parameter(Mandatory = $false, HelpMessage = 'Überspringt die Bestätigungsabfrage vor der Durchführung der Änderungen.')]
    [switch]$Force,

    [Parameter(Mandatory = $false, HelpMessage = 'Verwendet den "net user"-Befehl für den Passwort-Reset anstelle von Set-ADUser.')]
    [switch]$UseNetUser,

    [Parameter(Mandatory = $false, HelpMessage = 'Unterdrückt die ausführliche Konsolenausgabe (Verbose).')]
    [switch]$Silent
)

begin {
    #region Initialisierung und Lokalisierung
    # Detaillierte Informationen für das Skriptstart-Logging
    Write-Verbose "Skriptstart am $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')."
    Write-Verbose "PowerShell-Version: $($PSVersionTable.PSVersion)."
    Write-Verbose "Verwendete Parameter: OUNames='$($OUNames -join ', ')', ReportPath='$ReportPath', Force=$Force, UseNetUser=$UseNetUser, Silent=$Silent."

    # Standard-Fehlerbehandlung auf 'Stop' setzen, um Fehler sofort abzufangen
    $ErrorActionPreference = 'Stop'

    # Deutsche Lokalisierung einstellen für Datums-/Zeitformate und Zahlen
    Write-Verbose "Stelle Kultur auf 'de-DE' ein."
    $culture = [System.Globalization.CultureInfo]::GetCultureInfo('de-DE')
    [System.Threading.Thread]::CurrentThread.CurrentCulture = $culture
    [System.Threading.Thread]::CurrentThread.CurrentUICulture = $culture

    # Globale Liste für die Speicherung der Operationsergebnisse initialisieren
    Write-Verbose "Initialisiere globale Ergebnisliste 'OperationResults'."
    $global:OperationResults = [System.Collections.Generic.List[object]]::new()
    
    # Zeitpunkt des Passwort-Resets für das Logging festhalten
    $PasswordSetDate = Get-Date

    #region Hilfsfunktionen
    # Funktion zum Schreiben von Konsolen-Headern
    function Write-Header {
        param([string]$Title)
        if (-not $Silent) {
            Write-Host "`n$(('=' * 80))" -ForegroundColor Cyan
            Write-Host " $Title " -ForegroundColor Cyan
            Write-Host "$(('=' * 80))`n" -ForegroundColor Cyan
        }
        Write-Verbose "Header angezeigt: '$Title'."
    }

    # Funktion für benutzerdefinierte Write-Verbose Meldungen
    function Write-UserVerbose {
        param([string]$Message)
        if (-not $Silent) {
            Write-Verbose $Message
        }
    }
    #endregion Hilfsfunktionen

    # Parameter-Validierung und interaktive Abfrage
    if ($PSBoundParameters.ContainsKey('Silent')) {
        $VerbosePreference = 'SilentlyContinue'
    } else {
        $VerbosePreference = 'Continue'
    }

    # Interaktive Abfrage für OUNames, falls nicht via Parameter übergeben
    if (-not ($PSBoundParameters.ContainsKey('OUNames'))) {
        Write-Header -Title "INTERAKTIVE EINGABE: OU-NAMEN"
        $inputOUNames = Read-Host "Bitte geben Sie die OU-Namen durch Komma getrennt ein (z.B. 81, 82). Leer lassen für Standard: '$($OUNames -join ', ')'"
        if (-not [string]::IsNullOrWhiteSpace($inputOUNames)) {
            $OUNames = $inputOUNames.Split(',').Trim()
            Write-UserVerbose "OU-Namen interaktiv gesetzt: $($OUNames -join ', ')."
        } else {
            Write-UserVerbose "Verwende Standard-OU-Namen: $($OUNames -join ', ')."
        }
    }

    # Interaktive Abfrage für ReportPath, falls nicht via Parameter übergeben
    if (-not ($PSBoundParameters.ContainsKey('ReportPath'))) {
        Write-Header -Title "INTERAKTIVE EINGABE: BERICHTSPFAD"
        $inputReportPath = Read-Host "Bitte geben Sie den vollständigen Pfad für den Berichts-Export an. Leer lassen für Standard: '$ReportPath'"
        if (-not [string]::IsNullOrWhiteSpace($inputReportPath)) {
            $ReportPath = $inputReportPath.Trim()
            Write-UserVerbose "Berichtspfad interaktiv gesetzt: '$ReportPath'."
        } else {
            Write-UserVerbose "Verwende Standard-Berichtspfad: '$ReportPath'."
        }
    }
    #endregion Initialisierung und Lokalisierung
}

process {
    try {
        #region 1. OU-Suche
        Write-Header -Title "SCHRITT 1: ORGANISATIONSEINHEITEN-SUCHE"
        Write-Host "Suche nach OUs: $($OUNames -join ', ')" -ForegroundColor Yellow
        Write-UserVerbose "Beginne OU-Suche in der Domäne."

        $domain = Get-ADDomain | Select-Object -ExpandProperty DistinguishedName
        Write-UserVerbose "Domänen-DN ermittelt: '$domain'."

        $targetOUs = @()
        foreach ($name in $OUNames) {
            Write-UserVerbose "Suche nach OU mit Name '$name'."
            try {
                $ou = Get-ADOrganizationalUnit -Filter "Name -eq '$name'" `
                    -SearchBase $domain `
                    -SearchScope Subtree `
                    -ErrorAction Stop
                $targetOUs += $ou
                Write-UserVerbose "OU '$name' ($($ou.DistinguishedName)) gefunden."
            } catch {
                Write-Warning "OU '$name' konnte nicht gefunden werden: $($_.Exception.Message)"
                Write-Host "OU '$name' konnte nicht gefunden werden. Bitte überprüfen Sie den Namen." -ForegroundColor Yellow
            }
        }

        if (-not $targetOUs.Count) {
            Write-Host "Keine der angegebenen OUs ($($OUNames -join ', ')) konnte gefunden werden. Skript wird beendet." -ForegroundColor Red
            Write-UserVerbose "Keine OUs gefunden. Skript wird beendet."
            exit 1
        }
        Write-UserVerbose "$($targetOUs.Count) OUs gefunden."
        #endregion

        #region 2. Benutzersuche
        Write-Header -Title "SCHRITT 2: BENUTZERSUCHE"
        Write-Host "Suche nach Benutzern mit Namensmustern 'L110*' oder 'L114*'." -ForegroundColor Yellow
        Write-UserVerbose "Beginne Benutzersuche in den gefundenen OUs."

        $users = @()
        foreach ($ou in $targetOUs) {
            Write-UserVerbose "Suche Benutzer in OU: $($ou.DistinguishedName)."
            try {
                $foundUsers = Get-ADUser -LDAPFilter "(|(sAMAccountName=L110*)(sAMAccountName=L114*))" `
                    -SearchBase $ou.DistinguishedName `
                    -Properties Name, sAMAccountName, Enabled, LastLogonDate, DistinguishedName, `
                                PasswordNeverExpires, CannotChangePassword, PasswordLastSet, `
                                LastLogonTimestamp, UserAccountControl, PasswordChangeRequired `
                    -SearchScope Subtree `
                    -ErrorAction Stop
                $users += $foundUsers
                Write-UserVerbose "$($foundUsers.Count) Benutzer in $($ou.Name) gefunden."
            } catch {
                Write-Warning "Fehler beim Suchen von Benutzern in OU '$($ou.Name)': $($_.Exception.Message)"
            }
        }

        if (-not $users.Count) {
            Write-Host "Keine passenden Benutzer mit den Mustern 'L110*' oder 'L114*' in den angegebenen OUs gefunden." -ForegroundColor Yellow
            Write-UserVerbose "Keine passenden Benutzer gefunden. Skript wird beendet."
            return # Skript beenden, da keine Benutzer zum Verarbeiten vorhanden sind
        }
        Write-UserVerbose "$($users.Count) Benutzer insgesamt gefunden."
        #endregion

        #region 3. Benutzeranzeige
        Write-Header -Title "GEFUNDENE BENUTZER"
        Write-Host "Folgende Benutzer wurden zur Verarbeitung gefunden:" -ForegroundColor Green
        $users | Select-Object @{l='Benutzername';e={$_.Name}},
                                @{l='Anmeldename';e={$_.SamAccountName}},
                                @{l='Aktiviert';e={if($_.Enabled){'Ja'}else{'Nein'}}},
                                @{l='Passwort läuft nie ab';e={if($_.PasswordNeverExpires){'Ja'}else{'Nein'}}},
                                @{l='Passwort ändern nötig';e={if($_.PasswordChangeRequired){'Ja'}else{'Nein'}}},
                                @{l='Letzte Anmeldung';e={$_.LastLogonDate.ToString('dd.MM.yyyy HH:mm:ss', $culture)}},
                                @{l='OU-Pfad';e={$_.DistinguishedName}} | Format-Table -AutoSize
        Write-UserVerbose "Liste der gefundenen Benutzer in der Konsole angezeigt."
        #endregion

        #region 4. Passwort- und Bestätigungshinweis
        Write-Header -Title "PASSWORT- UND BESTÄTIGUNGSINFORMATION"
        Write-Host "Geplantes Passwort-Änderungsdatum für Protokollierung: $($PasswordSetDate.ToString('dd.MM.yyyy HH:mm:ss'))" -ForegroundColor Magenta
        Write-UserVerbose "Passwort-Änderungsdatum für Protokollierung gesetzt: $($PasswordSetDate.ToString('dd.MM.yyyy HH:mm:ss'))."

        if (-not $Force) {
            Write-UserVerbose "Warte auf Benutzerbestätigung, da '-Force' nicht gesetzt ist."
            $confirmation = Read-Host "Möchten Sie die Änderungen für die gefundenen Benutzer durchführen? (J/N)"
            if ($confirmation -notin @('J','j')) {
                Write-Host "Operation vom Benutzer abgebrochen." -ForegroundColor Yellow
                Write-UserVerbose "Operation vom Benutzer abgebrochen. Skript wird beendet."
                return # Skript beenden, da Benutzer abgebrochen hat
            }
            Write-UserVerbose "Benutzer hat die Änderungen bestätigt."
        }
        #endregion

        #region 5. Passwort-Eingabe
        Write-Header -Title "PASSWORT-EINGABE"
        Write-Host "Bitte geben Sie das neue Passwort für die betroffenen Benutzer ein." -ForegroundColor Green
        $securePass = Read-Host "Neues Passwort eingeben" -AsSecureString

        # Konvertierung des SecureString in einen Plaintext-String für 'net user'.
        # Dies ist nur notwendig, wenn 'net user' verwendet wird, und wird sofort danach bereinigt.
        $plainPass = $null
        if ($UseNetUser) {
            $plainPass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
                [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePass)
            )
            Write-UserVerbose "Passwort in Plaintext für 'net user' konvertiert."
        }
        #endregion

        #region 6. Benutzerverarbeitung (Passwort-Reset und Attribut-Update)
        Write-Header -Title "BENUTZERVERARBEITUNG: PASSWORT-RESET & ATTRBUIEREN"
        foreach ($user in $users) {
            # Template für das OperationResult-Objekt für jeden Benutzer
            $resultTemplate = [PSCustomObject]@{
                Benutzername        = $user.Name
                Anmeldename         = $user.SamAccountName
                Aktiviert           = $user.Enabled
                LetzteAnmeldung     = $user.LastLogonDate.ToString('dd.MM.yyyy HH:mm:ss', $culture) # Lokalisierte Ausgabe
                OU_Pfad             = $user.DistinguishedName
                PasswortNieAblauf   = $user.PasswordNeverExpires
                PasswortAendernNoetig = $user.PasswordChangeRequired
                OperationStatus     = "Nicht durchgeführt"
                FehlerCode          = $null
                FehlerMeldung       = $null
                PasswortDatum       = $null
            }
            Write-UserVerbose "Verarbeite Benutzer: $($user.SamAccountName)."

            try {
                if ($PSCmdlet.ShouldProcess($user.SamAccountName, "Passwort zurücksetzen und Attribute aktualisieren")) {
                    # --- Passwort-Reset und Attribut-Updates ---
                    if ($UseNetUser) {
                        # Option A: Passwort-Reset über 'net user'
                        Write-Host "`nVerarbeite Benutzer: $($user.SamAccountName) (mittels 'net user')" -ForegroundColor Cyan
                        Write-UserVerbose "Versuche Passwort-Reset für $($user.SamAccountName) mittels 'net user'."
                        $output = & net user $user.SamAccountName $plainPass /DOMAIN /ACTIVE:YES 2>&1
                        if ($LASTEXITCODE -ne 0) {
                            throw "NetUser-Fehler ($LASTEXITCODE): $output"
                        }
                        Write-UserVerbose "Passwort für $($user.SamAccountName) mittels 'net user' erfolgreich zurückgesetzt."

                        # AD-Einstellungen über Set-ADUser aktualisieren (Passwort-Attribute und LastLogonTimestamp)
                        Write-UserVerbose "Aktualisiere AD-Attribute für $($user.SamAccountName) mittels Set-ADUser."
                        Set-ADUser -Identity $user.DistinguishedName `
                                   -PasswordNeverExpires $true `
                                   -CannotChangePassword $true `
                                   -PasswordChangeRequired $false ` # Sicherstellen, dass dieses Flag nicht gesetzt ist
                                   -Replace @{lastLogonTimestamp = (Get-Date).ToFileTime()} ` # lastLogonTimestamp setzen
                                   -ErrorAction Stop
                        Write-UserVerbose "AD-Attribute für $($user.SamAccountName) erfolgreich aktualisiert."

                    } else {
                        # Option B: Passwort-Reset und Attribut-Update über 'Set-ADUser' (bevorzugt)
                        Write-Host "`nVerarbeite Benutzer: $($user.SamAccountName) (mittels 'Set-ADUser')" -ForegroundColor Cyan
                        Write-UserVerbose "Versuche Passwort-Reset und Attribut-Update für $($user.SamAccountName) mittels 'Set-ADUser'."
                        Set-ADUser -Identity $user.DistinguishedName `
                                   -NewPassword $securePass `
                                   -PasswordNeverExpires $true `
                                   -CannotChangePassword $true `
                                   -PasswordChangeRequired $false ` # Sicherstellen, dass dieses Flag nicht gesetzt ist
                                   -Replace @{lastLogonTimestamp = (Get-Date).ToFileTime()} ` # lastLogonTimestamp setzen
                                   -ErrorAction Stop
                        Write-UserVerbose "Passwort und AD-Attribute für $($user.SamAccountName) mittels 'Set-ADUser' erfolgreich aktualisiert."
                    }

                    # Erfolgsmeldung im Ergebnisobjekt aktualisieren
                    $resultTemplate.OperationStatus = "Erfolgreich"
                    $resultTemplate.PasswortDatum = (Get-Date).ToString('dd.MM.yyyy HH:mm:ss', $culture)
                    Write-Host "Benutzer '$($user.SamAccountName)' erfolgreich aktualisiert." -ForegroundColor Green
                    Write-UserVerbose "Benutzer '$($user.SamAccountName)' erfolgreich verarbeitet. Status: '$($resultTemplate.OperationStatus)'."
                } else {
                    $resultTemplate.OperationStatus = "Übersprungen (ShouldProcess)"
                    Write-Host "Verarbeitung von Benutzer '$($user.SamAccountName)' übersprungen (WhatIf/Confirm)." -ForegroundColor Yellow
                    Write-UserVerbose "Verarbeitung von Benutzer '$($user.SamAccountName)' übersprungen."
                }
            }
            catch {
                # Fehlerbehandlung für individuelle Benutzeroperationen
                $resultTemplate.OperationStatus = "Fehlgeschlagen"
                $resultTemplate.FehlerCode = $_.Exception.HResult # HResult kann spezifischere Fehlercodes liefern
                $resultTemplate.FehlerMeldung = $_.Exception.Message
                $resultTemplate.PasswortDatum = $null
                
                Write-Host "Fehler bei der Verarbeitung von Benutzer '$($user.SamAccountName)': $($_.Exception.Message)" -ForegroundColor Red
                Write-Warning "Fehler bei Benutzer '$($user.SamAccountName)': $($_.Exception.Message)"
                Write-UserVerbose "Fehler bei Benutzer '$($user.SamAccountName)'. Fehlermeldung: '$($_.Exception.Message)'."
            }
            finally {
                # Ergebnis des Benutzers zur globalen Liste hinzufügen
                $global:OperationResults.Add($resultTemplate)
                Write-UserVerbose "Ergebnis für Benutzer '$($user.SamAccountName)' zur Liste hinzugefügt."
            }
        }
        #endregion
    }
    catch {
        # Kritische Fehlerbehandlung für das gesamte Skript
        Write-Host "KRITISCHER FEHLER WÄHREND DER AUSFÜHRUNG: $($_.Exception.Message)" -ForegroundColor Red
        Write-Error "Kritischer Fehler: $($_.Exception.Message)"
        Write-UserVerbose "KRITISCHER FEHLER: $($_.Exception.Message)."
        exit 1 # Skript mit Fehlercode beenden
    }
    finally {
        #region 7. CSV-Export
        Write-Header -Title "CSV-EXPORT"
        if ($global:OperationResults.Count -gt 0) {
            try {
                # Export der Ergebnisse in eine CSV-Datei mit Semikolon als Trennzeichen und UTF8-Kodierung
                $global:OperationResults | Export-Csv -Path $ReportPath -Delimiter ";" -Encoding UTF8 -NoTypeInformation -ErrorAction Stop
                Write-Host "Bericht erfolgreich erstellt: $ReportPath" -ForegroundColor Green
                Write-UserVerbose "Bericht erfolgreich exportiert nach '$ReportPath'."
            } catch {
                Write-Host "Fehler beim Exportieren des Berichts nach '$ReportPath': $($_.Exception.Message)" -ForegroundColor Red
                Write-Warning "Export-Fehler: $($_.Exception.Message)"
                Write-UserVerbose "FEHLER: Bericht konnte nicht exportiert werden: '$($_.Exception.Message)'."
            }
        } else {
            Write-Host "Keine Ergebnisse zum Exportieren vorhanden." -ForegroundColor Yellow
            Write-UserVerbose "Keine Ergebnisse zum Exportieren vorhanden."
        }

        # Bereinigung des Plaintext-Passworts aus dem Speicher
        if ($plainPass) {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR(
                [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePass)
            )
            Remove-Variable plainPass -ErrorAction SilentlyContinue
            Write-UserVerbose "Plaintext-Passwort aus dem Speicher bereinigt."
        }
        Remove-Variable securePass -ErrorAction SilentlyContinue
        #endregion
    }
}

end {
    #region 8. Zusammenfassung
    Write-Header -Title "ZUSAMMENFASSUNG"
    $successCount = ($global:OperationResults | Where-Object { $_.OperationStatus -eq "Erfolgreich" }).Count
    $skippedCount = ($global:OperationResults | Where-Object { $_.OperationStatus -eq "Übersprungen (ShouldProcess)" }).Count
    $errorCount = ($global:OperationResults | Where-Object { $_.OperationStatus -eq "Fehlgeschlagen" }).Count

    Write-Host @"
    Zusammenfassung der Operationen:
    ================================
    Verarbeitete Benutzer: $($global:OperationResults.Count)
    Erfolgreich durchgeführt: $successCount
    Übersprungen:             $skippedCount
    Fehlgeschlagen:           $errorCount
"@ -ForegroundColor DarkCyan

    if ($errorCount -gt 0) {
        Write-Host "Detaillierte Fehlerübersicht:" -ForegroundColor Red
        $global:OperationResults | Where-Object { $_.OperationStatus -eq "Fehlgeschlagen" } |
            Format-Table Anmeldename, FehlerCode, FehlerMeldung -AutoSize | Out-String | Write-Host -ForegroundColor Red
        Write-UserVerbose "Skript beendet mit Fehlern. Fehleranzahl: $errorCount."
        exit 2 # Exit-Code für Fehler
    } else {
        Write-Host "Alle Operationen wurden erfolgreich durchgeführt oder übersprungen." -ForegroundColor Green
        Write-UserVerbose "Skript erfolgreich beendet. Keine Fehler aufgetreten."
        exit 0 # Exit-Code für Erfolg
    }
    #endregion
}
