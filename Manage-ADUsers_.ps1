<#
.SYNOPSIS
Ein Skript zur Verwaltung von Active Directory-Benutzern, einschließlich Passwort-Reset und Aktualisierung von Kontoeigenschaften mit erweiterter Protokollierung.

.DESCRIPTION
Dieses Skript automatisiert die Verwaltung von AD-Benutzern in spezifischen Organisationseinheiten (OUs).
Hauptfunktionen:
1. Suchen von OUs basierend auf Namen (z.B. '81', '82').
2. Finden von Benutzern in diesen OUs, deren Anmeldenamen mit 'L110' oder 'L114' beginnen.
3. Interaktive Anzeige der gefundenen Benutzer und Einholung einer Bestätigung vor der Durchführung von Änderungen.
4. Sicheres Zurücksetzen der Benutzerpasswörter.
5. Aktivieren der Konten, Setzen der Optionen "Passwort läuft nie ab" und "Benutzer kann Passwort nicht ändern".
6. Aktualisieren der Beschreibung des Benutzerkontos mit einem Zeitstempel der Änderung.
7. Detaillierte Protokollierung aller Aktionen in einer CSV-Datei mit deutscher Lokalisierung (Semikolon als Trennzeichen).
8. Unterstützt einen `-LogOnly` Modus für einen Testlauf ohne Änderungen und einen `-Force` Modus für nicht-interaktive Ausführung.

.PARAMETER OUNames
Ein Array von Zeichenketten, das die Namen der zu durchsuchenden OUs enthält.
Standardwert: @('81', '82')

.PARAMETER ReportPath
Der vollständige Pfad zur CSV-Datei, in der die Ergebnisse protokolliert werden.
Standardwert: "$env:TEMP\Benutzerbericht_JJJJMMTT_HHMMSS.csv"

.PARAMETER Force
Ein Schalter, der die interaktive Bestätigungsabfrage vor der Durchführung von Änderungen überspringt.

.PARAMETER LogOnly
Ein Schalter, um das Skript im reinen Protokollierungsmodus auszuführen. Es werden alle Benutzer gefunden und in der Konsole angezeigt, aber keine Änderungen vorgenommen. Nützlich für einen Testlauf.

.EXAMPLE
PS C:\> .\Manage-ADUsers_v3.1.ps1 -Verbose
Führt das Skript im interaktiven Modus mit detaillierter Ausgabe aus. Es sucht in den Standard-OUs '81' und '82', fragt nach einem neuen Passwort und bittet um Bestätigung, bevor Änderungen angewendet werden.

.EXAMPLE
PS C:\> .\Manage-ADUsers_v3.1.ps1 -OUNames 'Azubis', 'Praktikanten' -ReportPath "C:\ADLogs\Azubi_Reset.csv"
Sucht nach Benutzern in den OUs 'Azubis' und 'Praktikanten' und speichert den Bericht unter "C:\ADLogs\Azubi_Reset.csv".

.EXAMPLE
PS C:\> .\Manage-ADUsers_v3.1.ps1 -Force
Führt das Skript nicht-interaktiv aus. Ideal für geplante Aufgaben. Es wird direkt nach dem neuen Passwort gefragt und die Änderungen werden sofort angewendet.

.EXAMPLE
PS C:\> .\Manage-ADUsers_v3.1.ps1 -LogOnly
Führt eine Simulation durch. Das Skript findet und listet alle betroffenen Benutzer auf, nimmt aber keine Änderungen an den Konten vor.

.NOTES
Autor: PS-Coding (Überarbeitet basierend auf einer Vorlage)
Version: 3.1.20250921
Anforderungen: Das ActiveDirectory-Modul für PowerShell muss installiert und verfügbar sein. Das Skript muss in einem Kontext mit ausreichenden Berechtigungen zur Verwaltung von AD-Benutzern ausgeführt werden.

.CHANGELOG
v3.1 - 2025-09-21:
- [Sicherheit] Passwort-Reset erfolgt nun sicher über `Set-ADAccountPassword` mit einem `SecureString`. Die unsichere Konvertierung zu Klartext und die Verwendung von `net user` wurden entfernt.
- [Korrektheit] Das manuelle und fehleranfällige Setzen von `lastLogonTimestamp` wurde entfernt. Stattdessen wird die `Description`-Eigenschaft des Benutzers aktualisiert.
- [Best Practice] Ineffiziente Abfrage `Get-ADUser -Properties *` durch gezielte Abfrage der benötigten Eigenschaften ersetzt.
- [Best Practice] Globale Variable `$global:OperationResults` durch Skript-Variable `$script:OperationResults` ersetzt, um Konflikte zu vermeiden.
- [Robustheit] Das Skript prüft nun, ob der Zielordner für den Report existiert, und erstellt ihn bei Bedarf.
- [Funktionalität] Neuer Parameter `-LogOnly` für einen Testlauf ohne Änderungen hinzugefügt.
- [Benutzerfreundlichkeit] Ausgaben wurden verbessert und verwenden nun `Write-Verbose` für detaillierte Statusmeldungen.
- [Benutzerfreundlichkeit] Standard-Speicherort für den Report in das temporäre Verzeichnis des Benutzers verlegt, um Berechtigungsprobleme zu vermeiden.
- [Dokumentation] Hilfe und Kommentare wurden umfassend erweitert.
- [Struktur] Code wurde zur besseren Lesbarkeit und Wartbarkeit neu strukturiert.

.LINK
https://docs.microsoft.com/en-us/powershell/module/activedirectory/
#>

#Requires -Version 5.1
#Requires -Modules ActiveDirectory

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(HelpMessage = "Ein Array mit den Namen der zu durchsuchenden OUs.")]
    [string[]]$OUNames = @('81', '82'),

    [Parameter(HelpMessage = "Der vollständige Pfad für den CSV-Bericht.")]
    [string]$ReportPath = "$($env:TEMP)\Benutzerbericht_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",

    [Parameter(HelpMessage = "Überspringt die manuelle Bestätigung vor der Ausführung.")]
    [switch]$Force,

    [Parameter(HelpMessage = "Führt einen Testlauf durch, der Benutzer nur auflistet, aber keine Änderungen vornimmt.")]
    [switch]$LogOnly
)

begin {
    # Deutsche Kultur für Datums- und Zahlenformate festlegen
    $culture = [System.Globalization.CultureInfo]::GetCultureInfo('de-DE')
    [System.Threading.Thread]::CurrentThread.CurrentCulture = $culture
    [System.Threading.Thread]::CurrentThread.CurrentUICulture = $culture

    # Initialisierungen
    $ErrorActionPreference = 'Stop'
    $script:OperationResults = [System.Collections.Generic.List[object]]::new()
    $scriptExecutionTime = Get-Date

    # Helper-Funktion für formatierte Ausgaben (nutzt Verbose-Stream)
    function Write-Header {
        param([string]$Title)
        Write-Verbose "$(('=' * 80))"
        Write-Verbose " $Title "
        Write-Verbose "$(('=' * 80))"
    }

    Write-Verbose "Skriptausführung gestartet am $($scriptExecutionTime.ToString('dd.MM.yyyy HH:mm:ss'))"
}

process {
    try {
        # 1. OU-SUCHE
        Write-Header -Title "SCHRITT 1: ORGANISATIONSEINHEITEN-SUCHE"
        Write-Verbose "Suche nach OUs mit den Namen: $($OUNames -join ', ')"
        
        $domain = Get-ADDomain
        $targetOUs = foreach ($name in $OUNames) {
            # Fehlerbehandlung, falls eine bestimmte OU nicht gefunden wird
            try {
                Get-ADOrganizationalUnit -Filter "Name -eq '$name'" -Server $domain.DNSRoot -ErrorAction Stop
            }
            catch {
                Write-Warning "OU mit dem Namen '$name' konnte nicht gefunden werden."
            }
        }

        if (-not $targetOUs) {
            throw "Keine der angegebenen OUs konnte im Active Directory gefunden werden: $($OUNames -join ', ')"
        }
        Write-Verbose "Gefundene OU-Pfade: $($targetOUs.DistinguishedName -join '; ')"

        # 2. BENUTZERSUCHE
        Write-Header -Title "SCHRITT 2: BENUTZERSUCHE"
        $userProperties = @('Name', 'SamAccountName', 'Enabled', 'LastLogonDate', 'DistinguishedName', 'Description')
        $users = foreach ($ou in $targetOUs) {
            Write-Verbose "Durchsuche OU: $($ou.DistinguishedName)"
            Get-ADUser -LDAPFilter "(|(sAMAccountName=L110*)(sAMAccountName=L114*))" `
                -SearchBase $ou.DistinguishedName `
                -Properties $userProperties `
                -SearchScope Subtree
        }

        if (-not $users) {
            Write-Host "In den durchsuchten OUs wurden keine passenden Benutzer gefunden." -ForegroundColor Yellow
            return
        }

        # 3. BENUTZERANZEIGE UND VORAB-PROTOKOLLIERUNG
        Write-Header -Title "GEFUNDENE BENUTZER ($($users.Count))"
        $users | Format-Table @{l='Benutzername';e={$_.Name}},
                               @{l='Anmeldename';e={$_.SamAccountName}},
                               @{l='Aktiviert';e={if($_.Enabled){'Ja'}else{'Nein'}}},
                               @{l='Letzte Anmeldung';e={$_.LastLogonDate}} -AutoSize
        
        # 4. SIMULATIONSMODUS ODER BESTÄTIGUNG
        if ($LogOnly) {
            Write-Host "`n[INFO] Skript im LogOnly-Modus. Es werden keine Änderungen durchgeführt." -ForegroundColor Cyan
            # Benutzer trotzdem für den Report erfassen
            foreach ($user in $users) {
                $logEntry = [PSCustomObject]@{
                    Benutzername      = $user.Name
                    Anmeldename       = $user.SamAccountName
                    Aktiviert_vorher  = $user.Enabled
                    LetzteAnmeldung   = if($user.LastLogonDate) { $user.LastLogonDate } else { $null }
                    OU_Pfad           = $user.DistinguishedName
                    OperationStatus   = "LogOnly-Modus"
                    FehlerMeldung     = $null
                    Aenderungsdatum   = $null
                }
                $script:OperationResults.Add($logEntry)
            }
            return # Process-Block hier beenden
        }
        
        Write-Header -Title "BESTÄTIGUNG ERFORDERLICH"
        if (-not $Force) {
            $confirmation = Read-Host "Es wurden $($users.Count) Benutzer gefunden. Möchten Sie die Konten jetzt bearbeiten? (J/N)"
            if ($confirmation -notin @('J', 'j', 'Y', 'y')) {
                Write-Host "Aktion vom Benutzer abgebrochen." -ForegroundColor Yellow
                return
            }
        }

        # 5. SICHERE PASSWORTEINGABE
        Write-Header -Title "PASSWORTEINGABE"
        $securePass = Read-Host "Bitte geben Sie das neue Passwort für die Benutzer ein" -AsSecureString
        if ($securePass.Length -eq 0) {
            throw "Das Passwort darf nicht leer sein. Vorgang abgebrochen."
        }
        
        # 6. BENUTZERVERARBEITUNG
        Write-Header -Title "BENUTZERKONTEN WERDEN VERARBEITET"
        foreach ($user in $users) {
            $logEntry = [PSCustomObject]@{
                Benutzername      = $user.Name
                Anmeldename       = $user.SamAccountName
                Aktiviert_vorher  = $user.Enabled
                LetzteAnmeldung   = if($user.LastLogonDate) { $user.LastLogonDate } else { $null }
                OU_Pfad           = $user.DistinguishedName
                OperationStatus   = "Nicht durchgeführt"
                FehlerMeldung     = $null
                Aenderungsdatum   = $null
            }

            try {
                Write-Host "`nVerarbeite Benutzer: $($user.SamAccountName)" -ForegroundColor Cyan

                if ($PSCmdlet.ShouldProcess($user.DistinguishedName, "Konto aktivieren, Passwort zurücksetzen und Eigenschaften anpassen")) {
                    # Schritt 6.1: Konto aktivieren
                    Write-Verbose "Aktiviere Konto..."
                    Enable-ADAccount -Identity $user

                    # Schritt 6.2: Passwort sicher zurücksetzen
                    Write-Verbose "Setze Passwort zurück..."
                    Set-ADAccountPassword -Identity $user -NewPassword $securePass -Reset

                    # Schritt 6.3: Weitere AD-Eigenschaften anpassen
                    Write-Verbose "Setze Kontoeigenschaften..."
                    $newDescription = "Konto aktualisiert am $($scriptExecutionTime.ToString('dd.MM.yyyy HH:mm')) durch Skript."
                    Set-ADUser -Identity $user -PasswordNeverExpires $true -CannotChangePassword $true -Description $newDescription

                    # Erfolgreich verarbeitet
                    $logEntry.OperationStatus = "Erfolgreich"
                    $logEntry.Aenderungsdatum = Get-Date
                    Write-Host "Benutzer $($user.SamAccountName) erfolgreich aktualisiert." -ForegroundColor Green
                }
                else {
                    $logEntry.OperationStatus = "Übersprungen (-WhatIf)"
                    Write-Host "Benutzer $($user.SamAccountName) wurde aufgrund von -WhatIf übersprungen." -ForegroundColor Yellow
                }
            }
            catch {
                $logEntry.OperationStatus = "Fehlgeschlagen"
                $logEntry.FehlerMeldung = $_.Exception.Message.Trim()
                Write-Warning "Fehler bei der Verarbeitung von $($user.SamAccountName): $($_.Exception.Message)"
            }
            finally {
                $script:OperationResults.Add($logEntry)
            }
        }
    }
    catch {
        Write-Error "Ein kritischer Fehler ist im Skript aufgetreten: $($_.Exception.Message)"
        exit 1
    }
    finally {
        # 7. CSV-EXPORT
        if ($script:OperationResults.Count -gt 0) {
            Write-Header -Title "CSV-EXPORT"
            try {
                $ReportDir = Split-Path -Path $ReportPath -Parent
                if (-not (Test-Path -Path $ReportDir)) {
                    Write-Verbose "Erstelle Report-Verzeichnis: $ReportDir"
                    New-Item -ItemType Directory -Path $ReportDir -Force | Out-Null
                }
                
                $script:OperationResults | Export-Csv -Path $ReportPath -Delimiter ";" -Encoding UTF8 -NoTypeInformation
                Write-Host "`nBericht wurde erfolgreich erstellt: $ReportPath" -ForegroundColor Green
            }
            catch {
                Write-Warning "Der Bericht konnte nicht unter '$ReportPath' gespeichert werden: $($_.Exception.Message)"
            }
        }
    }
}

end {
    # 8. ZUSAMMENFASSUNG
    Write-Host "`n$(('=' * 80))" -ForegroundColor Cyan
    Write-Host " ZUSAMMENFASSUNG DER AUSFÜHRUNG" -ForegroundColor Cyan
    Write-Host "$(('=' * 80))`n" -ForegroundColor Cyan

    if ($script:OperationResults.Count -gt 0) {
        $successCount = ($script:OperationResults | Where-Object { $_.OperationStatus -eq "Erfolgreich" }).Count
        $errorCount = ($script:OperationResults | Where-Object { $_.OperationStatus -eq "Fehlgeschlagen" }).Count
        $skippedCount = ($script:OperationResults | Where-Object { $_.OperationStatus -like "*bersprungen*" -or $_.OperationStatus -like "*LogOnly*" }).Count

        Write-Host " Gesamt verarbeitete Benutzer: $($script:OperationResults.Count)"
        Write-Host " Erfolgreich:                 $successCount" -ForegroundColor Green
        Write-Host " Fehlgeschlagen:              $errorCount" -ForegroundColor Red
        Write-Host " Übersprungen/LogOnly:        $skippedCount" -ForegroundColor Yellow

        $failedOperations = $script:OperationResults | Where-Object { $_.OperationStatus -eq "Fehlgeschlagen" }
        if ($failedOperations) {
            Write-Host "`n--- Fehlerübersicht ---"
            $failedOperations | Format-Table Anmeldename, FehlerMeldung -AutoSize -Wrap | Out-String | Write-Host
        }
    }
    else {
        Write-Host "Es wurden keine Aktionen durchgeführt."
    }

    Write-Verbose "Skriptausführung beendet."
    if ($errorCount -gt 0) { exit 2 } # Exit-Code 2 für teilweise erfolgreiche Ausführung
}
