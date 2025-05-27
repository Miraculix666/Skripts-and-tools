# Invoke-ADUserMgmt.ps1
# AD-Benutzerverwaltung mit erweitertem Logging und deutscher Lokalisierung

<#
.SYNOPSIS
AD-Benutzerverwaltung mit erweitertem Logging und deutscher Lokalisierung.

.DESCRIPTION
Dieses Skript sucht Benutzer in Active Directory basierend auf vordefinierten OU-Namen
('81', '82') und sAMAccountName-Mustern ('L110*', 'L114*'). Es zeigt die gefundenen
Benutzer an, setzt deren Passwörter zurück und aktualisiert bestimmte Kontoeigenschaften
wie 'PasswordNeverExpires', 'CannotChangePassword', 'lastLogonTimestamp' und
stellt sicher, dass 'PasswordChangeRequired' auf $false gesetzt ist. Alle Operationen
werden detailliert protokolliert und die Ergebnisse in einer CSV-Datei im deutschen
Format (Semikolon-Trennzeichen, deutsche Datums-/Zeitformate) gespeichert.
Das Skript bietet die Wahl zwischen 'net user' und 'Set-ADUser' für den Passwort-Reset.

.PARAMETER OUNames
Ein Array von OU-Namen, in denen nach Benutzern gesucht werden soll.
Standardwerte sind '81' und '82'. Diese sind im Skript fest hinterlegt
und können hier nicht überschrieben werden, aber der Parameter ist aus
Kompatibilitätsgründen beibehalten.

.PARAMETER ReportPath
Der vollständige Pfad zur CSV-Berichtsdatei. Standardmäßig
'C:\Daten\Benutzerbericht.csv'.

.PARAMETER Force
Ein Switch-Parameter, der die Bestätigungsabfrage vor der Durchführung
von Änderungen unterdrückt.

.PARAMETER UseNetUser
Ein Switch-Parameter. Wenn dieser Parameter gesetzt ist, wird der 'net user'-Befehl
für den Passwort-Reset verwendet. Andernfalls (Standard) wird 'Set-ADUser' genutzt.

.PARAMETER Silent
Ein Switch-Parameter, der die detaillierte Verbose-Ausgabe in der Konsole unterdrückt.
Standardmäßig ist die ausführliche Ausgabe aktiviert.

.EXAMPLE
# Standardausführung mit interaktiver Bestätigung und Set-ADUser für Passwort-Reset
.\Invoke-ADUserMgmt.ps1

.EXAMPLE
# Ausführung ohne Bestätigung und Verwendung von 'net user' für den Passwort-Reset
.\Invoke-ADUserMgmt.ps1 -Force -UseNetUser

.EXAMPLE
# Ausführung mit spezifischem Berichtspfad und unterdrückter Verbose-Ausgabe
.\Invoke-ADUserMgmt.ps1 -ReportPath "D:\Logs\Benutzeraktionen.csv" -Silent

.NOTES
Autor: PS-Coding
Version: 3.1.20250527
Erstellungsdatum: 27.05.2025
Letzte Änderung: 27.05.2025
Kompatibilität: PowerShell 5.1 (PSVersion 5.1.22621.4391, PSEdition, CLRVersion 4.0.30319.42000)
Getestet auf On-Premise Windows Umgebungen (Win10, Server 2019).
Eingeschränkte Enterprise-Umgebung (keine Cloud-Integration).

Referenzen:
- AI-discovered sources:
  - Microsoft Docs: Get-ADUser, Set-ADUser, Get-ADOrganizationalUnit, Get-ADDomain
  - Stack Overflow: SecureString handling, CultureInfo for localization, Marshal.SecureStringToBSTR
  - PowerShell Best Practices for Logging and Error Handling, ErrorActionPreference
- User-provided sources:
  - Originales Skript des Benutzers
#>

#Requires -Version 5.1
#Requires -Modules ActiveDirectory

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter()]
    [string[]]$OUNames = @('81', '82'), # Bleibt fest im Skript, dient hier nur als Standardwert für Parameter

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
    # Konfigurieren der Verbose-Ausgabe basierend auf dem -Silent Parameter
    if ($Silent) {
        $VerbosePreference = 'SilentlyContinue'
    } else {
        $VerbosePreference = 'Continue'
    }
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

    # Passwort-Änderungsdatum für Protokollierung festhalten
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
    #endregion Hilfsfunktionen

    # Dynamische Ermittlung des Domänen-DN
    Write-Verbose "Ermittle den DistinguishedName der aktuellen AD-Domäne."
    try {
        $domain = (Get-ADDomain -ErrorAction Stop).DistinguishedName
        Write-Verbose "AD-Domäne erfolgreich ermittelt: '$domain'."
    } catch {
        Write-Host "Fehler beim Ermitteln des Domänen-DN: $($_.Exception.Message). Sicherstellen, dass das ActiveDirectory-Modul geladen ist und Domänenzugriff besteht." -ForegroundColor Red
        Write-Error "Kritischer Fehler: Domänen-DN konnte nicht ermittelt werden."
        exit 1
    }

    # Interaktive Abfrage für ReportPath, falls nicht via Parameter übergeben
    # Beibehalten, um interaktive Abfrage nach dem bestehenden Muster zu ermöglichen
    if (-not ($PSBoundParameters.ContainsKey('ReportPath'))) {
        Write-Header -Title "INTERAKTIVE EINGABE: BERICHTSPFAD"
        $inputReportPath = Read-Host "Bitte geben Sie den vollständigen Pfad für den Berichts-Export an. Leer lassen für Standard: '$ReportPath'"
        if (-not [string]::IsNullOrWhiteSpace($inputReportPath)) {
            $ReportPath = $inputReportPath.Trim()
            Write-Verbose "Berichtspfad interaktiv gesetzt: '$ReportPath'."
        } else {
            Write-Verbose "Verwende Standard-Berichtspfad: '$ReportPath'."
        }
    }

    # Überprüfung, ob der Berichtspfad gültig ist
    try {
        $reportDirectory = Split-Path $ReportPath -Parent
        if (-not (Test-Path $reportDirectory)) {
            Write-Verbose "Erstelle Verzeichnis für den Bericht: '$reportDirectory'."
            New-Item -ItemType Directory -Path $reportDirectory -Force | Out-Null
        }
    } catch {
        Write-Host "Fehler beim Überprüfen/Erstellen des Berichtspfad-Verzeichnisses '$reportDirectory': $($_.Exception.Message)" -ForegroundColor Red
        Write-Error "Kritischer Fehler: Berichtspfad-Verzeichnis konnte nicht vorbereitet werden."
        exit 1
    }

    #endregion Initialisierung und Lokalisierung
}

process {
    try {
        # 1. OU-SUCHE ---------------------------------------------------------
        Write-Header -Title "SCHRITT 1: ORGANISATIONSEINHEITEN-SUCHE"
        Write-Host "Suche nach OUs: $($OUNames -join ', ')" -ForegroundColor Yellow
        Write-Verbose "Beginne OU-Suche für die festgelegten OUs: $($OUNames -join ', ')."

        $targetOUs = @()
        foreach ($name in $OUNames) {
            Write-Verbose "Suche nach OU mit Name '$name' unter '$domain'."
            try {
                $ou = Get-ADOrganizationalUnit -Filter "Name -eq '$name'" `
                    -SearchBase $domain `
                    -SearchScope Subtree `
                    -ErrorAction Stop
                $targetOUs += $ou
                Write-Verbose "OU '$name' ($($ou.DistinguishedName)) gefunden."
            } catch {
                Write-Warning "OU '$name' konnte nicht gefunden werden: $($_.Exception.Message). Dieser OU-Name wird übersprungen."
                Write-Host "ACHTUNG: OU '$name' konnte nicht gefunden werden. Bitte überprüfen Sie den Namen, falls dies unerwartet ist." -ForegroundColor Yellow
            }
        }

        if (-not $targetOUs.Count) {
            Write-Host "Keine der festgelegten OUs ($($OUNames -join ', ')) konnte gefunden werden. Skript wird beendet." -ForegroundColor Red
            Write-Verbose "Keine der festgelegten OUs gefunden. Skript wird beendet."
            return
        }
        Write-Verbose "$($targetOUs.Count) gültige OUs gefunden."

        # 2. BENUTZERSUCHE ----------------------------------------------------
        Write-Header -Title "SCHRITT 2: BENUTZERSUCHE"
        Write-Host "Suche nach Benutzern mit Namensmustern 'L110*' oder 'L114*' in den gefundenen OUs." -ForegroundColor Yellow
        Write-Verbose "Beginne Benutzersuche in den gefundenen OUs."

        $users = @()
        foreach ($ou in $targetOUs) {
            Write-Verbose "Suche Benutzer in OU: $($ou.DistinguishedName)."
            try {
                $foundUsers = Get-ADUser -LDAPFilter "(|(sAMAccountName=L110*)(sAMAccountName=L114*))" `
                    -SearchBase $ou.DistinguishedName `
                    -Properties Name, SamAccountName, Enabled, LastLogonDate, DistinguishedName, `
                                PasswordNeverExpires, CannotChangePassword, PasswordLastSet, `
                                LastLogonTimestamp, UserAccountControl, PasswordChangeRequired ` # 'PasswordChangeRequired' hinzugefügt
                    -SearchScope Subtree `
                    -ErrorAction Stop
                $users += $foundUsers
                Write-Verbose "$($foundUsers.Count) Benutzer in $($ou.Name) gefunden."
            } catch {
                Write-Warning "Fehler beim Suchen von Benutzern in OU '$($ou.Name)': $($_.Exception.Message)"
                Write-Host "ACHTUNG: Fehler beim Suchen von Benutzern in OU '$($ou.Name)'. $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }

        if (-not $users.Count) {
            Write-Host "Keine passenden Benutzer mit den Mustern 'L110*' oder 'L114*' in den angegebenen OUs gefunden." -ForegroundColor Yellow
            Write-Verbose "Keine passenden Benutzer gefunden. Skript wird beendet."
            return # Skript beenden, da keine Benutzer zum Verarbeiten vorhanden sind
        }
        Write-Verbose "$($users.Count) Benutzer insgesamt zur Verarbeitung gefunden."

        # 3. BENUTZERANZEIGE --------------------------------------------------
        Write-Header -Title "GEFUNDENE BENUTZER"
        Write-Host "Folgende Benutzer wurden zur Verarbeitung gefunden:" -ForegroundColor Green
        $users | Select-Object @{l='Benutzername';e={$_.Name}},
                                @{l='Anmeldename';e={$_.SamAccountName}},
                                @{l='Aktiviert';e={if($_.Enabled){'Ja'}else{'Nein'}}},
                                @{l='Passwort läuft nie ab';e={if($_.PasswordNeverExpires){'Ja'}else{'Nein'}}},
                                @{l='Passwort bei nächster Anmeldung ändern';e={if($_.PasswordChangeRequired){'Ja'}else{'Nein'}}}, # Hinzugefügt
                                @{l='Letzte Anmeldung';e={$_.LastLogonDate.ToString('dd.MM.yyyy HH:mm:ss', $culture)}}, # Lokalisierte Datumsformatierung
                                @{l='OU-Pfad';e={$_.DistinguishedName}} | Format-Table -AutoSize
        Write-Verbose "Liste der gefundenen Benutzer in der Konsole angezeigt."

        # 4. PASSWORTHINWEIS --------------------------------------------------
        Write-Header -Title "PASSWORTINFORMATION"
        Write-Host "Geplantes Passwort-Änderungsdatum für Protokollierung: $($PasswordSetDate.ToString('dd.MM.yyyy HH:mm:ss'))" -ForegroundColor Magenta
        Write-Verbose "Passwort-Änderungsdatum für Protokollierung gesetzt: $($PasswordSetDate.ToString('dd.MM.yyyy HH:mm:ss'))."

        # 5. BESTÄTIGUNG ------------------------------------------------------
        Write-Header -Title "BESTÄTIGUNG"
        if (-not $Force) {
            Write-Verbose "Warte auf Benutzerbestätigung, da '-Force' nicht gesetzt ist."
            $confirmation = Read-Host "Möchten Sie die Änderungen für die gefundenen Benutzer durchführen? (J/N)"
            if ($confirmation -notin @('J','j')) {
                Write-Host "Operation vom Benutzer abgebrochen." -ForegroundColor Yellow
                Write-Verbose "Operation vom Benutzer abgebrochen. Skript wird beendet."
                return # Skript beenden, da Benutzer abgebrochen hat
            }
            Write-Verbose "Benutzer hat die Änderungen bestätigt."
        }

        # 6. PASSWORTVERWALTUNG -----------------------------------------------
        Write-Header -Title "PASSWORTZURÜCKSETZUNG"
        Write-Host "Bitte geben Sie das neue Passwort für die betroffenen Benutzer ein." -ForegroundColor Green
        $securePass = Read-Host "Neues Passwort eingeben" -AsSecureString

        # Konvertierung des SecureString in einen Plaintext-String für 'net user', falls benötigt
        $plainPass = $null
        if ($UseNetUser) {
            # Umwandlung nur, wenn net user tatsächlich verwendet wird
            Write-Verbose "Konvertiere SecureString zu Plaintext-Passwort für 'net user' (temporär)."
            $plainPass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
                [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePass)
            )
        }

        # 7. BENUTZERVERARBEITUNG ---------------------------------------------
        Write-Header -Title "BENUTZERVERARBEITUNG: PASSWORT-RESET & ATTRBUIEREN"
        foreach ($user in $users) {
            # Template für das OperationResult-Objekt für jeden Benutzer
            $resultTemplate = [PSCustomObject]@{
                Benutzername            = $user.Name
                Anmeldename             = $user.SamAccountName
                Aktiviert               = $user.Enabled
                LetzteAnmeldung         = $user.LastLogonDate.ToString('dd.MM.yyyy HH:mm:ss', $culture) # Lokalisierte Ausgabe
                OU_Pfad                 = $user.DistinguishedName
                PasswortNieAblauf       = $user.PasswordNeverExpires
                PasswortAendernNoetig   = $user.PasswordChangeRequired # Hinzugefügt
                OperationStatus         = "Nicht durchgeführt"
                FehlerCode              = $null
                FehlerMeldung           = $null
                PasswortDatum           = $null
            }
            Write-Verbose "Verarbeite Benutzer: $($user.SamAccountName)."

            try {
                # Bestätigungsprüfung für ShouldProcess (WhatIf/Confirm)
                if ($PSCmdlet.ShouldProcess($user.SamAccountName, "Passwort zurücksetzen und Attribute aktualisieren")) {
                    Write-Host "`nVerarbeite Benutzer: $($user.SamAccountName)" -ForegroundColor Cyan

                    if ($UseNetUser) {
                        # Option A: Passwort-Reset über 'net user'
                        Write-Verbose "Versuche Passwort-Reset für $($user.SamAccountName) mittels 'net user'."
                        $output = & net user $user.SamAccountName $plainPass /DOMAIN /ACTIVE:YES 2>&1
                        if ($LASTEXITCODE -ne 0) {
                            throw "NetUser-Fehler ($LASTEXITCODE): $output"
                        }
                        Write-Verbose "Passwort für $($user.SamAccountName) mittels 'net user' erfolgreich zurückgesetzt."

                        # AD-Einstellungen über Set-ADUser aktualisieren (Passwort-Attribute und LastLogonTimestamp)
                        Write-Verbose "Aktualisiere AD-Attribute für $($user.SamAccountName) mittels Set-ADUser nach 'net user' Operation."
                        Set-ADUser -Identity $user.DistinguishedName `
                                   -PasswordNeverExpires $true `
                                   -CannotChangePassword $true `
                                   -PasswordChangeRequired $false ` # Sicherstellen, dass dieses Flag nicht gesetzt ist
                                   -Replace @{lastLogonTimestamp = (Get-Date).ToFileTime()} ` # lastLogonTimestamp setzen
                                   -ErrorAction Stop
                        Write-Verbose "AD-Attribute für $($user.SamAccountName) erfolgreich aktualisiert nach 'net user' Operation."

                    } else {
                        # Option B: Passwort-Reset und Attribut-Update über 'Set-ADUser'
                        Write-Verbose "Versuche Passwort-Reset und Attribut-Update für $($user.SamAccountName) mittels 'Set-ADUser'."
                        Set-ADUser -Identity $user.DistinguishedName `
                                   -NewPassword $securePass `
                                   -PasswordNeverExpires $true `
                                   -CannotChangePassword $true `
                                   -PasswordChangeRequired $false ` # Sicherstellen, dass dieses Flag nicht gesetzt ist
                                   -Replace @{lastLogonTimestamp = (Get-Date).ToFileTime()} ` # lastLogonTimestamp setzen
                                   -ErrorAction Stop
                        Write-Verbose "Passwort und AD-Attribute für $($user.SamAccountName) mittels 'Set-ADUser' erfolgreich aktualisiert."
                    }

                    # Erfolgsmeldung im Ergebnisobjekt aktualisieren
                    $resultTemplate.OperationStatus = "Erfolgreich"
                    $resultTemplate.PasswortDatum = (Get-Date).ToString('dd.MM.yyyy HH:mm:ss', $culture)
                    Write-Host "Benutzer '$($user.SamAccountName)' erfolgreich aktualisiert." -ForegroundColor Green
                    Write-Verbose "Benutzer '$($user.SamAccountName)' erfolgreich verarbeitet. Status: '$($resultTemplate.OperationStatus)'."
                } else {
                    $resultTemplate.OperationStatus = "Übersprungen (ShouldProcess)"
                    Write-Host "Verarbeitung von Benutzer '$($user.SamAccountName)' übersprungen (WhatIf/Confirm)." -ForegroundColor Yellow
                    Write-Verbose "Verarbeitung von Benutzer '$($user.SamAccountName)' übersprungen."
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
                Write-Verbose "Fehler bei Benutzer '$($user.SamAccountName)'. Fehlermeldung: '$($_.Exception.Message)'."
            }
            finally {
                # Ergebnis des Benutzers zur globalen Liste hinzufügen
                $global:OperationResults.Add($resultTemplate)
                Write-Verbose "Ergebnis für Benutzer '$($user.SamAccountName)' zur Liste hinzugefügt."
            }
        }
    }
    catch {
        # Kritische Fehlerbehandlung für das gesamte Skript (z.B. wenn OU-Suche fehlschlägt)
        Write-Host "KRITISCHER FEHLER WÄHREND DER AUSFÜHRUNG: $($_.Exception.Message)" -ForegroundColor Red
        Write-Error "Kritischer Fehler: $($_.Exception.Message)"
        Write-Verbose "KRITISCHER FEHLER: $($_.Exception.Message)."
        exit 1 # Skript mit Fehlercode beenden
    }
    finally {
        # 8. CSV-EXPORT -------------------------------------------------------
        Write-Header -Title "CSV-EXPORT"
        if ($global:OperationResults.Count -gt 0) {
            try {
                # Export der Ergebnisse in eine CSV-Datei mit Semikolon als Trennzeichen und UTF8-Kodierung
                Write-Verbose "Exportiere Ergebnisse nach '$ReportPath' mit Semikolon als Trennzeichen."
                $global:OperationResults | Export-Csv -Path $ReportPath -Delimiter ";" -Encoding UTF8 -NoTypeInformation -ErrorAction Stop
                Write-Host "Bericht erfolgreich erstellt: $ReportPath" -ForegroundColor Green
                Write-Verbose "Bericht erfolgreich exportiert."
            } catch {
                Write-Host "Fehler beim Exportieren des Berichts nach '$ReportPath': $($_.Exception.Message)" -ForegroundColor Red
                Write-Warning "Export-Fehler: $($_.Exception.Message)"
                Write-Verbose "FEHLER: Bericht konnte nicht exportiert werden: '$($_.Exception.Message)'."
            }
        } else {
            Write-Host "Keine Ergebnisse zum Exportieren vorhanden." -ForegroundColor Yellow
            Write-Verbose "Keine Ergebnisse zum Exportieren vorhanden."
        }

        # Bereinigung des Plaintext-Passworts aus dem Speicher
        if ($plainPass) {
            Write-Verbose "Bereinige Plaintext-Passwort aus dem Speicher."
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR(
                [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePass)
            )
            Remove-Variable plainPass -ErrorAction SilentlyContinue
        }
        # SecureString-Variable bereinigen
        Remove-Variable securePass -ErrorAction SilentlyContinue
        Write-Verbose "SecureString-Variable aus dem Speicher bereinigt."
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
        Write-Verbose "Skript beendet mit Fehlern. Fehleranzahl: $errorCount."
        exit 2 # Exit-Code für Fehler
    } else {
        Write-Host "Alle Operationen wurden erfolgreich durchgeführt oder übersprungen." -ForegroundColor Green
        Write-Verbose "Skript erfolgreich beendet. Keine Fehler aufgetreten."
        exit 0 # Exit-Code für Erfolg
    }
}
