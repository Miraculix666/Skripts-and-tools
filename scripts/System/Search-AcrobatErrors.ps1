<#
.SYNOPSIS
    Sucht nach Adobe Acrobat-Fehlern im Windows-Anwendungsprotokoll und speichert die Ergebnisse.

.DESCRIPTION
    Dieses Skript durchsucht das Windows-Anwendungsprotokoll nach Ereignissen, die von verschiedenen Adobe Acrobat-bezogenen
    Providern stammen und den Ereignislevel 'Fehler' (Level 2) oder 'Kritisch' (Level 1) haben.
    Es bietet sowohl einen interaktiven Modus (ohne Parameter) als auch einen parametrisierten Modus.
    Die Ausgabe wird formatiert und in eine Logdatei geschrieben, deren Name Datum und Uhrzeit enthält.
    Das Skript ist für die deutsche Lokalisierung optimiert und bietet detaillierte, optionale Ausgaben.

.PARAMETER LogFilePath
    Der vollständige Pfad zur Logdatei. Wenn nicht angegeben, wird ein Standardpfad im Skriptverzeichnis verwendet.
    Der Dateiname enthält automatisch ein Datum und eine Uhrzeit.

.PARAMETER LogLevel
    Der Schwellenwert für den Ereignislevel, der im Log gesucht werden soll.
    Standardmäßig werden 'Fehler' (2) und 'Kritisch' (1) gesucht.
    Mögliche Werte: 1 (Kritisch), 2 (Fehler), 3 (Warnung), 4 (Information), 5 (Überwachungserfolg), 0 (Alle).
    Kann als Array angegeben werden (z.B. -LogLevel 1,2,3).

.PARAMETER ProviderName
    Ein oder mehrere Namen der Ereignisanbieter, nach denen gesucht werden soll.
    Standardmäßig werden gängige Adobe Acrobat-bezogene Provider verwendet.
    Beispiel: -ProviderName 'Adobe Acrobat', 'Acrobat Reader'.

.PARAMETER SuppressVerbose
    Verhindert die Ausgabe detaillierter Informationen auf der Konsole (Verbose-Ausgabe).
    Standardmäßig ist die Verbose-Ausgabe aktiviert.

.PARAMETER Keywords
    Eine Liste von Schlüsselwörtern, nach denen in den Ereignismeldungen gesucht werden soll, falls keine spezifischen
    Adobe-Provider gefunden werden. Standardmäßig werden 'Adobe' und 'Acrobat' verwendet.
    Diese Option wird nur verwendet, wenn keine Providernamen identifiziert werden können.

.EXAMPLE
    .\Find-AdobeAcrobatErrors.ps1
    Startet das Skript im interaktiven Modus und fragt nach den Einstellungen.

.EXAMPLE
    .\Find-AdobeAcrobatErrors.ps1 -LogLevel 1 -ProviderName 'Adobe Acrobat'
    Sucht nur nach kritischen Fehlern von 'Adobe Acrobat' und schreibt diese ins Log.

.EXAMPLE
    .\Find-AdobeAcrobatErrors.ps1 -SuppressVerbose
    Sucht mit Standardeinstellungen, unterdrückt aber die detaillierte Ausgabe auf der Konsole.

.EXAMPLE
    .\Find-AdobeAcrobatErrors.ps1 -LogFilePath "C:\Logs\AcrobatErrors.log" -LogLevel 1,2 -ProviderName 'Adobe Reader'
    Sucht nach kritischen Fehlern und Fehlern von 'Adobe Reader' und speichert sie in einer spezifischen Datei.

.EXAMPLE
    .\Find-AdobeAcrobatErrors.ps1 -Keywords 'PDF', 'ReaderError' -LogLevel 2
    Sucht nach Fehlern im Anwendungsprotokoll, deren Nachrichten 'PDF' oder 'ReaderError' enthalten, wenn keine Adobe-Provider gefunden werden.

.NOTES
    Autor: PS-Coding (Optimiert)
    Datum: 2025-07-03 (Letzte Anpassung: 2025-07-10)

    Quellen:
    - PowerShell Get-WinEvent Cmdlet Documentation - https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.diagnostics/get-winevent
    - PowerShell Write-Verbose, Write-Error, Write-Host Cmdlets - https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/write-verbose
    - PowerShell Parameter attributes and advanced functions - https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_functions_advanced
    - Ursprüngliches Skript für Adobe Acrobat Fehlerprüfung

    Bitte stellen Sie sicher, dass Sie über die entsprechenden Berechtigungen zum Lesen der Ereignisprotokolle verfügen.
#>

#File: Find-AdobeAcrobatErrors.ps1
#Description: Ein PowerShell-Skript zum Suchen und Protokollieren von Adobe Acrobat-Fehlern im Windows-Anwendungsprotokoll.

[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'Interactive')]
param(
    [Parameter(Mandatory = $false, HelpMessage = "Der vollständige Pfad zur Logdatei.")]
    [string]$LogFilePath,

    [Parameter(Mandatory = $false, HelpMessage = "Der Schwellenwert für den Ereignislevel (1=Kritisch, 2=Fehler, 3=Warnung, 4=Information, 5=Überwachungserfolg, 0=Alle).")]
    [ValidateSet(0, 1, 2, 3, 4, 5)] # 0=Alle, 1=Kritisch, 2=Fehler, 3=Warnung, 4=Information, 5=Überwachungserfolg
    [int[]]$LogLevel = @(1, 2), # Standard: Kritisch und Fehler

    [Parameter(Mandatory = $false, HelpMessage = "Ein oder mehrere Namen der Ereignisanbieter, nach denen gesucht werden soll.")]
    [string[]]$ProviderName = @('Adobe Acrobat', 'Acrobat Reader', 'AdobeARM', 'Adobe Acrobat Update Service'), # Gängige Acrobat-bezogene Provider

    [Parameter(Mandatory = $false, HelpMessage = "Unterdrückt die detaillierte Ausgabe auf der Konsole.")]
    [switch]$SuppressVerbose,

    [Parameter(Mandatory = $false, HelpMessage = "Schlüsselwörter, nach denen in Ereignismeldungen gesucht werden soll, falls keine Adobe-Provider gefunden werden.")]
    [string[]]$Keywords = @('Adobe', 'Acrobat', 'PDF') # Default keywords for message search
)

# Region: Variablen- und Parameterdefinitionen
# -------------------------------------------

# Speicherort des Skripts
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Standard-Dateiname für das Log (Datum und Uhrzeit im Namen, Deutsch formatiert)
$DefaultLogFileName = "AdobeAcrobat_Fehler_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Standard-Logpfad, wenn nicht anders angegeben
if (-not $LogFilePath) {
    $LogFilePath = Join-Path -Path $ScriptPath -ChildPath $DefaultLogFileName
}

# Globale Variable für die Verbose-Ausgabe
$Global:VerboseOutputEnabled = -not $SuppressVerbose

# Region: Hilfsfunktionen
# ---------------------

function Write-MyVerbose {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )
    if ($Global:VerboseOutputEnabled) {
        Write-Verbose $Message -Verbose
    }
}

function Write-MyHost {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [string]$ForegroundColor = 'White'
    )
    Write-Host $Message -ForegroundColor $ForegroundColor
    Add-LogEntry -Message $Message -Level 'INFO' # Auch in die Logdatei schreiben
}

function Add-LogEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [string]$Level = 'INFO' # INFO, WARN, ERROR, EVENT (for actual event data)
    )

    # Sicherstellen, dass die Logdatei existiert oder erstellt wird
    try {
        $logDir = Split-Path -Path $LogFilePath
        if (-not (Test-Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -ErrorAction Stop | Out-Null
            Write-MyVerbose "Log-Verzeichnis '$logDir' wurde erstellt."
        }
    }
    catch {
        Write-Error "Fehler beim Erstellen des Log-Verzeichnisses '$logDir': $($_.Exception.Message)"
        return # Abbruch, wenn Verzeichnis nicht erstellt werden kann
    }

    # Format für Datum und Uhrzeit (Deutsch)
    $culture = [System.Globalization.CultureInfo]::GetCultureInfo("de-DE")
    $timestamp = (Get-Date).ToString("dd.MM.yyyy HH:mm:ss", $culture)

    $logMessage = "[$timestamp] [$Level] $Message"
    try {
        $logMessage | Out-File -FilePath $LogFilePath -Append -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        Write-Error "Fehler beim Schreiben in die Logdatei '$LogFilePath': $($_.Exception.Message)"
    }
}

# Region: Interaktiver Modus
# -------------------------

function Handle-InteractiveMode {
    Write-MyHost "--- Adobe Acrobat Fehler Suche (Interaktiver Modus) ---" -ForegroundColor Cyan
    Write-MyHost "Bitte geben Sie die gewünschten Einstellungen ein oder drücken Sie Enter für den Standardwert." -ForegroundColor Yellow

    # LogFilePath abfragen
    if (-not ($PSBoundParameters.ContainsKey('LogFilePath'))) {
        $inputPath = Read-Host "Geben Sie den Pfad für die Logdatei ein (Standard: '$LogFilePath')"
        if (-not [string]::IsNullOrWhiteSpace($inputPath)) {
            $script:LogFilePath = $inputPath
        }
    }
    Write-MyVerbose "Verwendeter Logpfad: '$script:LogFilePath'"

    # LogLevel abfragen
    if (-not ($PSBoundParameters.ContainsKey('LogLevel'))) {
        $defaultLogLevelString = ($LogLevel -join ',')
        $inputLevel = Read-Host "Geben Sie die gewünschten Ereignislevel ein (z.B. 1,2 für Kritisch,Fehler; 0 für Alle. Standard: $defaultLogLevelString)"
        if (-not [string]::IsNullOrWhiteSpace($inputLevel)) {
            try {
                $levelArray = $inputLevel.Split(',') | ForEach-Object { [int]([string]$_).Trim() }
                # Validate each parsed integer against the valid set
                foreach ($level in $levelArray) {
                    if ($level -notin (0, 1, 2, 3, 4, 5)) {
                        throw "Ungültiger Levelwert '$level'. Gültige Werte sind 0, 1, 2, 3, 4, 5."
                    }
                }
                $script:LogLevel = $levelArray
            }
            catch {
                Write-Error "Ungültige Eingabe für LogLevel: $($_.Exception.Message). Verwende Standardwerte ($defaultLogLevelString)." `
                -ErrorAction Continue
                $script:LogLevel = @(1, 2) # Ensure default is set if invalid input
            }
        }
    }
    Write-MyVerbose "Verwendete Log-Level: $($script:LogLevel -join ', ')"

    # ProviderName abfragen
    if (-not ($PSBoundParameters.ContainsKey('ProviderName'))) {
        $defaultProviderString = ($ProviderName -join ', ')
        $inputProvider = Read-Host "Geben Sie die gewünschten Provider-Namen ein (durch Kommas getrennt. Standard: $defaultProviderString)"
        if (-not [string]::IsNullOrWhiteSpace($inputProvider)) {
            $script:ProviderName = $inputProvider.Split(',') | ForEach-Object { ([string]$_).Trim() }
        }
    }
    Write-MyVerbose "Verwendete Provider: $($script:ProviderName -join ', ')"

    # Keywords abfragen (if not provided as parameter)
    if (-not ($PSBoundParameters.ContainsKey('Keywords'))) {
        $defaultKeywordsString = ($Keywords -join ', ')
        $inputKeywords = Read-Host "Geben Sie Schlüsselwörter für die Nachrichtensuche ein (durch Kommas getrennt. Standard: $defaultKeywordsString)"
        if (-not [string]::IsNullOrWhiteSpace($inputKeywords)) {
            $script:Keywords = $inputKeywords.Split(',') | ForEach-Object { ([string]$_).Trim() }
        }
    }
    Write-MyVerbose "Verwendete Schlüsselwörter für Nachrichtensuche: $($script:Keywords -join ', ')"

    # SuppressVerbose abfragen
    if (-not ($PSBoundParameters.ContainsKey('SuppressVerbose'))) {
        $inputVerbose = Read-Host "Möchten Sie die detaillierte Konsolenausgabe unterdrücken? (J/N, Standard: N)"
        if ($inputVerbose -eq 'J' -or $inputVerbose -eq 'j') {
            $script:SuppressVerbose = $true
            $Global:VerboseOutputEnabled = $false
        } else {
            $script:SuppressVerbose = $false
            $Global:VerboseOutputEnabled = $true
        }
    }
    Write-MyVerbose "Detaillierte Konsolenausgabe aktiviert: $Global:VerboseOutputEnabled"
}

# Region: Hauptlogik
# ------------------

# Prüfen, ob der interaktive Modus ausgeführt werden soll
if ($PSCmdlet.ParameterSetName -eq 'Interactive' -and $PSBoundParameters.Count -eq 0) {
    Write-MyVerbose "Interaktiver Modus (keine Parameter übergeben) erkannt. Abfrage der Parameter."
    Handle-InteractiveMode
} elseif ($PSCmdlet.ParameterSetName -eq 'Interactive' -and $PSBoundParameters.Count -gt 0) {
    Write-MyVerbose "Parameterisierter Modus erkannt. Überprüfe und fülle fehlende Parameter interaktiv auf."
    Handle-InteractiveMode
} else {
    Write-MyVerbose "Parameterisierter Modus erkannt. Verwende bereitgestellte Parameter und Standardwerte."
}

# Prä-Prozessierung: Zusammenfassung der Einstellungen und Bestätigung
Write-MyHost "--- Zusammenfassung der Einstellungen ---" -ForegroundColor Cyan
Write-MyHost "Ziel-Logdatei: '$LogFilePath'" -ForegroundColor White
Write-MyHost "Gesuchte Ereignis-Level: $($LogLevel -join ', ')" -ForegroundColor White
Write-MyHost "Gesuchte Provider: $($ProviderName -join ', ')" -ForegroundColor White
Write-MyHost "Gesuchte Schlüsselwörter (falls Provider nicht gefunden): $($Keywords -join ', ')" -ForegroundColor White
Write-MyHost "Detaillierte Konsolenausgabe: $(if ($Global:VerboseOutputEnabled) {'Aktiviert'} else {'Deaktiviert'})" -ForegroundColor White
Write-MyHost "---------------------------------------" -ForegroundColor Cyan

# Bestätigung vor der Ausführung
if ($PSCmdlet.ShouldProcess("Adobe Acrobat Fehler suchen und protokollieren", "Soll die Suche mit diesen Einstellungen gestartet werden?")) {
    Write-MyHost "Suche nach Adobe Acrobat Fehlern im Anwendungsprotokoll..." -ForegroundColor Green
    Add-LogEntry -Message "Skriptstart - Suche nach Adobe Acrobat Fehlern." -Level 'INFO'
    Add-LogEntry -Message "Parameter: LogFilePath='$LogFilePath', LogLevel='$($LogLevel -join ', ')', ProviderName='$($ProviderName -join ', ')'" -Level 'INFO'
    Add-LogEntry -Message "Parameter: Keywords='$($Keywords -join ', ')'" -Level 'INFO'

    try {
        Write-MyVerbose "Starte Get-WinEvent-Abfrage..."

        # Validate ProviderNames first
        $validProviders = @()
        if ($ProviderName.Count -gt 0) {
            foreach ($provider in $ProviderName) {
                try {
                    $providerInfo = Get-WinEvent -ListProvider $provider -ErrorAction SilentlyContinue
                    if ($providerInfo) {
                        $validProviders += $provider
                        Write-MyVerbose "Provider '$provider' wurde gefunden und wird verwendet."
                    } else {
                        Write-MyHost "WARNUNG: Provider '$provider' wurde nicht gefunden. Er wird übersprungen." -ForegroundColor Yellow
                        Add-LogEntry -Message "WARNUNG: Provider '$provider' wurde nicht gefunden. Er wird übersprungen." -Level 'WARN'
                    }
                } catch {
                    Write-MyHost "WARNUNG: Fehler beim Prüfen von Provider '$provider': $($_.Exception.Message). Er wird übersprungen." -ForegroundColor Yellow
                    Add-LogEntry -Message "WARNUNG: Fehler beim Prüfen von Provider '$provider': $($_.Exception.Message). Er wird übersprungen." -Level 'WARN'
                }
            }
        }

        $searchByKeywords = $false
        if (-not $validProviders.Count) {
            Write-MyHost "Es wurden keine der angegebenen Adobe-bezogenen Provider gefunden." -ForegroundColor Red
            Write-MyHost "Möchten Sie stattdessen eine Suche nach allen 'Adobe'-Providern versuchen, die auf diesem System existieren? (J/N)" -ForegroundColor Yellow
            $confirmAutoDiscover = Read-Host "Bitte geben Sie 'J' für Ja oder 'N' für Nein ein"

            if ($confirmAutoDiscover -eq 'J' -or $confirmAutoDiscover -eq 'j') {
                Write-MyHost "Suche nach allen 'Adobe'-Providern auf diesem System..." -ForegroundColor Cyan
                try {
                    $discoveredProviders = (Get-WinEvent -ListProvider *Adobe* -ErrorAction SilentlyContinue).Name
                    if ($discoveredProviders.Count -gt 0) {
                        $script:ProviderName = $discoveredProviders
                        Write-MyHost "Folgende Adobe-Provider wurden automatisch gefunden und werden verwendet: $($script:ProviderName -join ', ')" -ForegroundColor Green
                        Add-LogEntry -Message "Automatisch gefundene Adobe-Provider: $($script:ProviderName -join ', ')" -Level 'INFO'
                        $validProviders = @($script:ProviderName) # Update validProviders with discovered ones
                    } else {
                        Write-MyHost "Es konnten keine Adobe-Provider automatisch auf diesem System gefunden werden." -ForegroundColor Yellow
                        Add-LogEntry -Message "INFO: Keine Adobe-Provider automatisch gefunden. Versuche Suche nach Schlüsselwörtern." -Level 'INFO'
                        $searchByKeywords = $true # Fallback to keyword search
                    }
                } catch {
                    Write-MyHost "WARNUNG: Fehler bei der automatischen Suche nach Adobe-Providern: $($_.Exception.Message). Versuche Suche nach Schlüsselwörtern." -ForegroundColor Yellow
                    Add-LogEntry -Message "WARNUNG: Fehler bei der automatischen Provider-Suche: $($_.Exception.Message). Versuche Suche nach Schlüsselwörtern." -Level 'WARN'
                    $searchByKeywords = $true # Fallback to keyword search
                }
            } else {
                Write-MyHost "Keine Providernamen bestätigt. Skript wird Suche nach Schlüsselwörtern vorschlagen." -ForegroundColor Yellow
                Add-LogEntry -Message "INFO: Keine Providernamen bestätigt. Versuche Suche nach Schlüsselwörtern." -Level 'INFO'
                $searchByKeywords = $true
            }
        }

        # Determine the final filtering approach
        $finalFilterLogName = 'Application'
        $finalFilterProviderName = @()
        $finalFilterKeywords = @()

        if ($validProviders.Count -gt 0) {
            $finalFilterProviderName = $validProviders
            Write-MyVerbose "Suche wird mit den validierten Providern durchgeführt: $($finalFilterProviderName -join ', ')"
        } elseif ($searchByKeywords -and $Keywords.Count -gt 0) {
            Write-MyHost "Führe Suche nach Schlüsselwörtern in Ereignismeldungen durch, da keine Provider gefunden wurden." -ForegroundColor Cyan
            Add-LogEntry -Message "Führe Suche nach Schlüsselwörtern in Ereignismeldungen durch, da keine Provider gefunden wurden." -Level 'INFO'
            $finalFilterKeywords = $Keywords
            # When searching by keywords, we get all entries first then filter
            Write-MyVerbose "Suche wird mit Schlüsselwörtern: $($finalFilterKeywords -join ', ') durchgeführt."
        } else {
            Write-Error "Weder gültige Provider noch Schlüsselwörter für die Suche verfügbar. Skript kann nicht fortgesetzt werden." -ErrorAction Stop
            Write-MyHost "Skript abgebrochen." -ForegroundColor Red
            Add-LogEntry -Message "FEHLER: Weder gültige Provider noch Schlüsselwörter für die Suche verfügbar. Skript abgebrochen." -Level 'ERROR'
            return
        }

        # Construct the initial filter based on whether we have providers or will search by keywords
        $initialFilter = @{
            LogName = $finalFilterLogName
        }
        if ($finalFilterProviderName.Count -gt 0) {
            $initialFilter.ProviderName = $finalFilterProviderName
        }

        # Retrieve events with initial filter.
        $rawEvents = Get-WinEvent -FilterHashtable $initialFilter -ErrorAction SilentlyContinue

        $events = $null # Initialize $events to null

        # Post-process the events based on LogLevel AND Keywords if applicable
        if ($LogLevel -contains 0) {
            Write-MyVerbose "LogLevel enthält '0' (Alle Ebenen). Alle gefundenen Ereignisse werden vorab verarbeitet."
            $events = $rawEvents
        } else {
            Write-MyVerbose "Filtere Ereignisse nach den angegebenen Log-Leveln: $($LogLevel -join ', ')"
            $events = $rawEvents | Where-Object { $_.Level -in $LogLevel }
        }

        # If we are in keyword search mode, filter the messages
        if ($finalFilterKeywords.Count -gt 0) {
            Write-MyVerbose "Führe zusätzliche Filterung nach Schlüsselwörtern in den Ereignismeldungen durch."
            $keywordPattern = ($finalFilterKeywords | ForEach-Object { [regex]::Escape($_) }) -join '|'
            $events = $events | Where-Object { $_.Message -match $keywordPattern }
        }


        if ($events) {
            Write-MyVerbose "Gefundene Ereignisse nach allen Filterungen: $($events.Count)"
            Write-MyVerbose "Verarbeite und schreibe Ereignisse in '$LogFilePath'..."
            Add-LogEntry -Message "Anzahl der gefundenen Ereignisse nach Filterung: $($events.Count)" -Level 'INFO'

            # Header für die Logdatei hinzufügen
            Add-LogEntry -Message "--- Start der Adobe Acrobat Fehlerprotokollierung ---" -Level 'INFO'
            Add-LogEntry -Message "Timestamp (DD.MM.YYYY HH:MM:SS) ; Quelle ; Ereignis-ID ; Nachricht" -Level 'INFO'

            $events | ForEach-Object {
                # Deutsche Datums- und Zeitformatierung
                $culture = [System.Globalization.CultureInfo]::GetCultureInfo("de-DE")
                $Timestamp = $_.TimeCreated.ToString("dd.MM.yyyy HH:mm:ss", $culture)
                $Message = $_.Message -replace "`r`n", " " # Zeilenumbrüche entfernen für bessere CSV/Log-Lesbarkeit
                $FormattedLine = "$Timestamp ; Quelle: $($_.ProviderName) ; Ereignis-ID: $($_.Id) ; Nachricht: $Message"

                # Direkter Aufruf von Add-LogEntry, um in die Datei zu schreiben und verbosely auszugeben
                Add-LogEntry -Message $FormattedLine -Level 'EVENT' # Verwende 'EVENT' für tatsächliche Ereignisdaten
                Write-MyVerbose "Protokolliert: $($FormattedLine.Substring(0, [System.Math]::Min(120, $FormattedLine.Length)))" # Kürzen für Verbose-Ausgabe
            }

            Write-MyHost "Suche abgeschlossen. Ergebnisse erfolgreich in '$LogFilePath' geschrieben." -ForegroundColor Green
            Add-LogEntry -Message "Skriptende - Suche abgeschlossen. Ergebnisse in '$LogFilePath' geschrieben." -Level 'INFO'
        }
        else {
            Write-MyHost "Keine passenden Adobe Acrobat Ereignisse im Anwendungsprotokoll gefunden." -ForegroundColor Yellow
            Add-LogEntry -Message "Keine passenden Adobe Acrobat Ereignisse gefunden." -Level 'INFO'
        }
    }
    catch {
        Write-Error "Ein schwerwiegender Fehler ist aufgetreten: $($_.Exception.Message)" -ErrorAction Continue
        Write-MyHost "Die Fehler konnten nicht vollständig verarbeitet und in die Datei geschrieben werden." -ForegroundColor Red
        Add-LogEntry -Message "Fehler: $($_.Exception.Message)" -Level 'ERROR'
        Add-LogEntry -Message "StackTrace: $($_.ScriptStackTrace)" -Level 'ERROR'
    }
}
else {
    Write-MyHost "Skriptausführung wurde vom Benutzer abgebrochen." -ForegroundColor Red
    Add-LogEntry -Message "Skriptausführung vom Benutzer abgebrochen." -Level 'INFO'
}

# --- ENDE DES SKRIPTS ---
