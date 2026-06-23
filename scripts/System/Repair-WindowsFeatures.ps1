<#
.SYNOPSIS
    Führt eine Reparaturinstallation eines optionalen Windows-Features durch.
    Deaktiviert und reaktiviert das angegebene Windows-Feature (z.B. "WindowsMediaPlayer").

.DESCRIPTION
    Dieses Skript überprüft das Vorhandensein des angegebenen Windows-Features.
    Falls es aktiviert ist, wird es deaktiviert und anschließend wieder aktiviert.
    Dies entspricht einer Reparaturinstallation. Das Skript ist für die Verwendung
    in einer On-Premise-Umgebung optimiert und berücksichtigt deutsche Regionaleinstellungen.

.PARAMETER FeatureName
    Der Name des Windows-Features, das repariert werden soll. Standard: "WindowsMediaPlayer".

.PARAMETER Force
    Führt die Reparatur ohne eine interaktive Bestätigungsaufforderung durch.

.PARAMETER Quiet
    Unterdrückt die Anzeige von Standard- und Verbose-Ausgaben.

.PARAMETER Verbose
    Aktiviert die detaillierte Ausgabe, die den Fortschritt des Skripts nachverfolgt.
    Diese Option ist standardmäßig aktiviert, kann aber mit -Quiet deaktiviert werden.

.EXAMPLE
    .\Repair-WindowsFeatures.ps1 -FeatureName "WindowsMediaPlayer"
    Führt die WMP-Reparatur im interaktiven Modus aus und fordert zur Bestätigung auf.

.EXAMPLE
    .\Repair-WindowsFeatures.ps1 -FeatureName "WindowsMediaPlayer" -Force
    Führt die WMP-Reparaturinstallation ohne Bestätigung aus.

.NOTES
    Autor: PS-Coding
    Version: 1.1
    Datum: 26.08.2025

    Entwicklungsquellen:
    - Microsoft Docs: Get-WindowsOptionalFeature, Enable-WindowsOptionalFeature, Disable-WindowsOptionalFeature
    - Community-Diskussionen zur Windows-Feature-Verwaltung
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false, HelpMessage = "Name des optionalen Windows-Features, das repariert werden soll.")]
    [string]$FeatureName = "WindowsMediaPlayer",

    [switch]$Force,
    [switch]$Quiet,
    [switch]$Verbose
)

#-------------------------------------------------------------------------------------------------------------------
# PS-Coding-Standardfunktionen
#-------------------------------------------------------------------------------------------------------------------

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = 'INFO'
    )
    if (-not $Quiet) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Write-Host "[$timestamp] [$Level]: $Message"
    }
}

#-------------------------------------------------------------------------------------------------------------------
# Pre-Processing Instructions
#-------------------------------------------------------------------------------------------------------------------
# Setze die deutsche Kultur für die korrekte Datums- und Zahlenformatierung.
try {
    Write-Log -Message "Setze die Kultur auf 'de-DE' für deutsche Regionaleinstellungen." -Level "INFO"
    [System.Threading.Thread]::CurrentThread.CurrentCulture = 'de-DE'
    [System.Threading.Thread]::CurrentThread.CurrentUICulture = 'de-DE'
}
catch {
    Write-Log -Message "Fehler beim Setzen der Kultur auf 'de-DE': $($_.Exception.Message)" -Level "WARN"
}

# Standardmäßig Verbose-Ausgabe aktivieren, falls nicht explizit unterdrückt.
if (-not $Quiet) {
    $VerbosePreference = "Continue"
} else {
    $VerbosePreference = "SilentlyContinue"
}

#-------------------------------------------------------------------------------------------------------------------
# Skriptlogik
#-------------------------------------------------------------------------------------------------------------------

Write-Log -Message "Starte die Reparaturinstallation für '$FeatureName'." -Level "INFO"

# Überprüfe, ob das Feature vorhanden ist.
try {
    Write-Verbose "Überprüfe den Status des optionalen Features '$FeatureName'..."
    $featureStatus = Get-WindowsOptionalFeature -Online -FeatureName $FeatureName -ErrorAction Stop
}
catch {
    Write-Log -Message "Fehler beim Abrufen des Features '$FeatureName'. Möglicherweise ist es nicht verfügbar. Fehler: $($_.Exception.Message)" -Level "ERROR"
    return
}

# Zeige den aktuellen Status an.
Write-Log -Message "Aktueller Status von '$FeatureName': $($featureStatus.State)" -Level "INFO"

# Führe eine Bestätigungsaufforderung durch, falls der Force-Parameter nicht gesetzt ist.
if (-not $Force) {
    $confirmationMessage = "Möchten Sie wirklich eine Reparaturinstallation für '$FeatureName' durchführen?"
    if (-not $PSCmdlet.ShouldContinue($confirmationMessage, "Bestätigung erforderlich")) {
        Write-Log -Message "Reparatur vom Benutzer abgebrochen." -Level "INFO"
        return
    }
}

# Deaktiviere das Feature zur Reparatur.
if ($featureStatus.State -eq 'Enabled') {
    Write-Log -Message "Deaktiviere '$FeatureName' zur Reparatur..." -Level "INFO"
    try {
        Disable-WindowsOptionalFeature -Online -FeatureName $FeatureName -NoRestart -ErrorAction Stop
        Write-Verbose "'$FeatureName' erfolgreich deaktiviert."
    }
    catch {
        Write-Log -Message "Fehler beim Deaktivieren von '$FeatureName': $($_.Exception.Message)" -Level "ERROR"
        return
    }
} else {
    Write-Log -Message "'$FeatureName' ist bereits deaktiviert. Führe direkt die Aktivierung durch." -Level "INFO"
}

# Aktiviere das Feature, um die Reparatur abzuschließen.
Write-Log -Message "Aktiviere '$FeatureName' zur Reparatur..." -Level "INFO"
try {
    Enable-WindowsOptionalFeature -Online -FeatureName $FeatureName -NoRestart -ErrorAction Stop
    Write-Log -Message "Reparaturinstallation für '$FeatureName' erfolgreich abgeschlossen." -Level "SUCCESS"
}
catch {
    Write-Log -Message "Fehler beim Aktivieren von '$FeatureName': $($_.Exception.Message)" -Level "ERROR"
}

# Verifizierung nach der Installation.
Write-Verbose "Überprüfe den Status nach der Reparatur..."
$finalStatus = Get-WindowsOptionalFeature -Online -FeatureName $FeatureName
if ($finalStatus.State -eq 'Enabled') {
    Write-Log -Message "'$FeatureName' ist nach der Reparatur erfolgreich aktiviert." -Level "SUCCESS"
} else {
    Write-Log -Message "Überprüfung fehlgeschlagen. '$FeatureName' konnte nicht aktiviert werden. Aktueller Status: $($finalStatus.State)" -Level "ERROR"
}
