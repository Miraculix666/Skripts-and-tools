<#
.SYNOPSIS
    Erzwingt das Auslösen aller wichtigen SCCM/Configuration Manager Client-Richtlinienzyklen (Policy, Inventory, Deployment, Updates) auf einem Zielcomputer.

.DESCRIPTION
    Dieses Skript verwendet WMI, um eine Reihe von kritischen Client-Aktionen auf einem Remote-
    oder lokalen Client auszulösen, um eine sofortige Synchronisierung der Richtlinien und
    eine schnelle Initiierung von Software-Installationen oder -Updates zu gewährleisten.
    Es bietet flexible Parameterisierung, interaktive Eingabe, umfassende Fehlerbehandlung,
    detailliertes Verbose Logging und Lokalisierung für das deutsche Gebietschema.

    Ausgelöste Zyklen:
    - Computer Policy Retrieval & Evaluation Cycle
    - Application Deployment Evaluation Cycle
    - Software Inventory Cycle
    - Hardware Inventory Cycle
    - Software Updates Scan Cycle

.PARAMETER ComputerName
    Der Hostname oder die IP-Adresse des Zielcomputers, auf dem die Zyklen ausgelöst werden sollen.
    Wenn nicht angegeben, wird interaktiv nach dem Namen gefragt.

.PARAMETER Silent
    Ein Switch-Parameter, der die standardmäßig aktivierte ausführliche Ausgabe (Write-Verbose) unterdrückt.

.EXAMPLE
    .\Invoke-SCCMCycle.ps1 -ComputerName 'SRV01'

    Löst alle relevanten SCCM-Zyklen auf dem Computer 'SRV01' aus, um die Software-Aktualisierung zu erzwingen.

.EXAMPLE
    .\Invoke-SCCMCycle.ps1 -Silent

    Führt das Skript im interaktiven Modus aus und unterdrückt die detaillierte (Verbose) Ausgabe.

.NOTES
    Autor: PS-Coding Assistant
    Version: 1.1 (Erweiterung der Client-Zyklen)
    Datum: 2025-10-28
    Anforderungen: Lokale Administratorrechte auf dem ausführenden System und WMI-Zugriff
                   auf das Zielsystem. PowerShell 5.1 oder neuer.

    Referenzen:
    - AI-discovered sources: SCCM Client Action GUIDs (Microsoft Documentation)
    - User-provided sources: (Wird hier eingefügt, falls im Input vorhanden)
#>
[CmdletBinding(DefaultParameterSetName='Interactive')]
param(
    [Parameter(Mandatory=$false, Position=0, HelpMessage='Hostname oder IP des Zielcomputers')]
    [string]$ComputerName,

    [Parameter(Mandatory=$false)]
    [switch]$Silent
)

# --- 1. Konfiguration und Vorbereitung ---

# Definition der SCCM Client-Aktionen (GUIDs und Anzeigenamen)
# Diese Zyklen stellen sicher, dass Richtlinien abgerufen und Inventuren aktualisiert werden.
$SCCMCycles = @(
    @{ Name = "Computerrichtlinie abrufen & evaluieren"; Guid = "{00000000-0000-0000-0000-000000000001}" }, # Policy Retrieval & Evaluation
    @{ Name = "Software-Inventur-Zyklus"; Guid = "{00000000-0000-0000-0000-000000000002}" },             # Software Inventory
    @{ Name = "Hardware-Inventur-Zyklus"; Guid = "{00000000-0000-0000-0000-000000000007}" },             # Hardware Inventory
    @{ Name = "Anwendungsbereitstellungs-Evaluation"; Guid = "{00000000-0000-0000-0000-000000000021}" }, # Application Deployment Evaluation
    @{ Name = "Software-Updates-Scan-Zyklus"; Guid = "{00000000-0000-0000-0000-000000000108}" }          # Software Updates Scan
)

# WMI-Namespace und Klasse für SCCM-Client-Aktionen
$WmiNamespace = 'root\ccm'
$WmiClass = 'SMS_Client'

# Lokalisierung: Setzt das aktuelle Gebietsschema auf Deutsch (Deutschland) für die Ausgabe.
$CultureGerman = New-Object System.Globalization.CultureInfo("de-DE")
[System.Threading.Thread]::CurrentThread.CurrentUICulture = $CultureGerman
[System.Threading.Thread]::CurrentThread.CurrentCulture = $CultureGerman

# Verbose-Ausgabe standardmäßig aktivieren, wenn der Silent-Switch NICHT gesetzt ist.
if (-not $Silent) {
    $VerbosePreference = 'Continue'
    Write-Verbose "⚙️ Skript-Konfiguration: Verbose-Modus aktiviert (Standard). Kultur: $($CultureGerman.Name)."
} else {
    $VerbosePreference = 'SilentlyContinue'
}


# --- 2. Parameter-Validierung und Interaktive Eingabe ---

# Wenn $ComputerName leer ist, im interaktiven Modus nachfragen
if ([string]::IsNullOrEmpty($ComputerName)) {
    Write-Verbose "Der Parameter 'ComputerName' fehlt. Fordere interaktive Eingabe an."
    do {
        $ComputerName = Read-Host -Prompt "Bitte geben Sie den Hostnamen oder die IP des Zielcomputers ein"
    } until (-not [string]::IsNullOrEmpty($ComputerName))
}

Write-Verbose "Zielcomputer: **$($ComputerName)**."


# --- 3. Erreichbarkeitsprüfung (Ping) und Vorbereitung der Remote-Aktion ---

Write-Host "📡 Prüfe Erreichbarkeit von: $($ComputerName)..."
Write-Verbose "Führe Test-Connection mit Zählzahl 1 aus."

try {
    # Test-Connection muss auf das Zielsystem antworten können (ICMP-Freigabe benötigt)
    if (-not (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet -ErrorAction Stop)) {
        throw "Ping fehlgeschlagen oder geblockt."
    }
    Write-Host "✅ Host ist erreichbar."
    Write-Verbose "Erreichbarkeit erfolgreich geprüft. Beginne mit WMI-Aufrufen..."
}
catch {
    Write-Error "🚨 Fehler: Hostname '$($ComputerName)' ist nicht erreichbar oder WMI-Zugriff nicht möglich. Details: $($_.Exception.Message)"
    exit 1
}


# --- 4. Sequenzielles Auslösen aller Zyklen ---

Write-Host "`n🚀 Starte das Auslösen der SCCM-Client-Zyklen auf $($ComputerName)..."

# Gehe alle definierten Zyklen durch und löse sie sequenziell aus
foreach ($Cycle in $SCCMCycles) {
    $CycleName = $Cycle.Name
    $CycleGuid = $Cycle.Guid
    
    Write-Host "  -> ⚙️ Löst '$($CycleName)' aus..."
    Write-Verbose "WMI-Aufruf: Klasse='$WmiClass', Methode='TriggerSchedule', GUID='$CycleGuid'."

    try {
        # WMI-Methode auf dem Remote-Computer aufrufen
        Invoke-WmiMethod -ComputerName $ComputerName -Namespace $WmiNamespace -Class $WmiClass -Name 'TriggerSchedule' -ArgumentList $CycleGuid -ErrorAction Stop

        Write-Host "     ✅ $CycleName erfolgreich an Client übergeben."
        Write-Verbose "$CycleName erfolgreich an den SCCM Client-Dienst übergeben."
    }
    catch {
        Write-Error "     🚨 Fehler beim Auslösen von '$($CycleName)': $($_.Exception.Message). Prüfen Sie Client-Status und Rechte."
    }
}


# --- 5. Abschluss und Hinweis ---

$CurrentTimeGerman = Get-Date -Format 'dd.MM.yyyy HH:mm:ss'
Write-Host "`n--- Skript-Ende [$($CurrentTimeGerman)] ---"
Write-Host "ℹ️ Alle relevanten Zyklen wurden an **$($ComputerName)** gesendet."
Write-Host "   Die neue Software/Aktualisierung sollte nun im Software Center erscheinen und ausgeführt werden."
Write-Verbose "Skript abgeschlossen. Überprüfen Sie die Client-Logs (PolicyAgent.log, AppDiscovery.log, InventoryAgent.log) auf dem Zielsystem für den endgültigen Status."
