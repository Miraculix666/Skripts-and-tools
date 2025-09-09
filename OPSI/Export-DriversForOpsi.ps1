# FILENAME: Export-DriversForOpsi.ps1
#
# DESCRIPTION: This script exports Windows drivers to a specified OPSI depot directory.
#              It automatically determines the computer's manufacturer and model to create
#              a structured folder path (...\<Manufacturer>\<Model>). The script includes
#              features for parameterization, administrative privilege checking, interactive
#              confirmation, and verbose logging.
#
# SOURCES:
#   AI-discovered sources: N/A
#   User-provided sources: Initial script provided by user.
#

<#
.SYNOPSIS
    Exports Windows drivers to a structured directory path for an OPSI depot.

.DESCRIPTION
    This script is designed for Windows system administrators to easily backup local drivers
    to a network share, specifically formatted for an OPSI byAudit structure.

    It performs the following steps:
    1. Checks for Administrator privileges.
    2. Determines the computer manufacturer and model.
    3. Constructs a target UNC path based on the OPSI depot root, manufacturer, and model.
    4. Prompts the user for confirmation before proceeding.
    5. Creates the target directory if it does not exist.
    6. Executes the Export-WindowsDriver command.
    7. Verifies that drivers were successfully exported.

    The script can be run interactively or fully parameterized for automation.

.PARAMETER OpsiDepotPath
    Specifies the root path of the OPSI driver depot. This is the base path where the
    manufacturer and model subdirectories will be created.

.PARAMETER Manufacturer
    Overrides the automatic detection of the computer manufacturer.

.PARAMETER Model
    Overrides the automatic detection of the computer model.

.PARAMETER Force
    A switch to bypass the interactive confirmation prompt. Use with caution.

.EXAMPLE
    PS C:\> .\Export-DriversForOpsi.ps1

    Runs the script in interactive mode. It will prompt for the OPSI Depot Path.

.EXAMPLE
    PS C:\> .\Export-DriversForOpsi.ps1 -OpsiDepotPath "\\opsi-server\opsi_depot_rw\win10-x64\drivers\additional\byAudit"

    Exports drivers for the current machine to the specified depot path, after asking for confirmation.

.EXAMPLE
    PS C:\> .\Export-DriversForOpsi.ps1 -OpsiDepotPath "\\fileserver\depot\drivers" -Manufacturer "LENOVO" -Model "20XW004QGE" -Force

    Exports drivers for the current machine, overriding manufacturer and model, to a custom path
    and bypasses the confirmation prompt. This is useful for scripting or automation.

.EXAMPLE
    PS C:\> .\Export-DriversForOpsi.ps1 -Verbose:$false

    Runs the script without the detailed verbose output.

.NOTES
    Author: PS-Coding (AI Assistant)
    Version: 2.0
    Requires: PowerShell 5.1 or higher. Must be run as Administrator.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory = $false, HelpMessage = "Geben Sie den Stamm-Zielpfad für das OPSI-Treiberdepot an.")]
    [string]$OpsiDepotPath,

    [Parameter(Mandatory = $false, HelpMessage = "Überschreibt den automatisch erkannten Hersteller.")]
    [string]$Manufacturer,

    [Parameter(Mandatory = $false, HelpMessage = "Überschreibt das automatisch erkannte Modell.")]
    [string]$Model,

    [Parameter(Mandatory = $false, HelpMessage = "Überspringt die Bestätigungsabfrage vor der Ausführung.")]
    [switch]$Force
)

#================================================================================
# --- INITIALIZATION AND PRE-CHECKS ---
#================================================================================

# Standardmäßig ausführliche Ausgabe aktivieren. Kann mit -Verbose:$false unterdrückt werden.
$VerbosePreference = 'Continue'

Write-Verbose "Skript gestartet: Export-DriversForOpsi.ps1"

# 1. Administratorrechte prüfen
Write-Verbose "Prüfe auf Administratorrechte..."
if (-not ([System.Security.Principal.WindowsPrincipal][System.Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "Dieses Skript erfordert Administratorrechte. Bitte starten Sie PowerShell als Administrator und führen Sie das Skript erneut aus."
    # Das Skript wird hier beendet, da der Fehler "terminating" ist.
    exit 1
}
Write-Verbose "Administratorrechte bestätigt."

# Interaktiver Modus: Fehlende Parameter abfragen
if (-not $OpsiDepotPath) {
    Write-Host "INTERAKTIVER MODUS" -ForegroundColor Yellow
    $OpsiDepotPath = Read-Host "Bitte geben Sie den OPSI-Depot-Stammpfad ein (z.B. \\server\share\drivers)"
}

#================================================================================
# --- CORE LOGIC ---
#================================================================================

try {
    # 2. Hersteller und Modell auslesen, falls nicht als Parameter übergeben
    if (-not $Manufacturer -or -not $Model) {
        Write-Verbose "Hersteller und Modell werden vom lokalen System ausgelesen..."
        $computerInfo = Get-CimInstance -ClassName Win32_ComputerSystem
        # Parameter-Werte nur dann überschreiben, wenn sie nicht explizit gesetzt wurden
        if (-not $Manufacturer) { $Manufacturer = $computerInfo.Manufacturer }
        if (-not $Model) { $Model = $computerInfo.Model }
        Write-Verbose "Hersteller gefunden: '$Manufacturer'"
        Write-Verbose "Modell gefunden: '$Model'"
    }
    else {
        Write-Verbose "Hersteller und Modell aus Parametern verwendet: '$Manufacturer' / '$Model'"
    }


    # Ersetzt ungültige Zeichen in Ordnernamen, um robuste Pfade zu gewährleisten
    $hersteller_clean = $Manufacturer -replace '[\\/:*?"<>|]', '_'
    $modell_clean = $Model -replace '[\\/:*?"<>|]', '_'
    Write-Verbose "Bereinigter Herstellername für Ordner: '$hersteller_clean'"
    Write-Verbose "Bereinigter Modellname für Ordner: '$modell_clean'"


    # 3. Zielpfad für den Treiberexport zusammenbauen
    $zielpfad = Join-Path -Path $OpsiDepotPath -ChildPath (Join-Path -Path $hersteller_clean -ChildPath $modell_clean)
    Write-Verbose "Vollständiger Zielpfad wurde erstellt: '$zielpfad'"


    #================================================================================
    # --- PRE-PROCESSING & CONFIRMATION ---
    #================================================================================

    Write-Host "`nZUSAMMENFASSUNG DER AKTIONEN:" -ForegroundColor Cyan
    Write-Host "----------------------------------"
    Write-Host "Hersteller      : $Manufacturer ($hersteller_clean)"
    Write-Host "Modell          : $Model ($modell_clean)"
    Write-Host "Zielverzeichnis : $zielpfad"
    Write-Host "----------------------------------"

    # Bestätigung vom Benutzer einholen, falls -Force nicht verwendet wird
    if ($Force -or $PSCmdlet.ShouldProcess($zielpfad, "Treiber exportieren nach")) {
        if (-not $Force) {
            $confirmation = Read-Host "`nMöchten Sie den Treiberexport durchführen? (J/N)"
        }

        if ($Force -or $confirmation -match '^[jJ]$') {
            Write-Verbose "Benutzer hat die Aktion bestätigt oder -Force wurde verwendet."

            # 4. Zielverzeichnis erstellen, falls es nicht existiert
            Write-Verbose "Prüfe, ob das Zielverzeichnis existiert..."
            if (-not (Test-Path -Path $zielpfad)) {
                Write-Verbose "Zielverzeichnis existiert nicht. Es wird erstellt."
                New-Item -ItemType Directory -Path $zielpfad -Force | Out-Null
                Write-Host "Verzeichnis erfolgreich erstellt: $zielpfad" -ForegroundColor Green
            }
            else {
                Write-Verbose "Zielverzeichnis existiert bereits."
            }

            # 5. Treiber exportieren
            Write-Host "`nStarte den Export der Treiber. Dies kann einige Minuten dauern..." -ForegroundColor Yellow
            Export-WindowsDriver -Online -Destination $zielpfad
            Write-Verbose "Befehl 'Export-WindowsDriver' wurde abgeschlossen."

            #================================================================================
            # --- SUCCESS VERIFICATION ---
            #================================================================================
            Write-Verbose "Überprüfe den Erfolg des Exports..."
            $exportedFiles = Get-ChildItem -Path $zielpfad
            if ($exportedFiles.Count -gt 0) {
                Write-Host "`nERFOLG: $($exportedFiles.Count) Treiberpakete wurden erfolgreich nach '$zielpfad' exportiert." -ForegroundColor Green
            }
            else {
                Write-Warning "FEHLER: Der Export wurde zwar abgeschlossen, aber es wurden keine Treiber im Zielverzeichnis gefunden. Bitte überprüfen Sie die Berechtigungen und den verfügbaren Speicherplatz."
            }

        }
        else {
            Write-Host "`nAbbruch durch Benutzer." -ForegroundColor Red
        }
    }
    else {
        Write-Host "`nAbbruch durch Benutzer (ShouldProcess)." -ForegroundColor Red
    }

}
catch {
    # Fängt alle "terminating" Fehler im 'try' Block ab
    Write-Error "Ein schwerwiegender Fehler ist aufgetreten: $($_.Exception.Message)"
    Write-Error "Das Skript wird beendet."
}

# Das PowerShell Fenster bleibt am Ende offen für die Überprüfung der Ausgabe.
# Write-Host "`nDrücken Sie eine beliebige Taste zum Beenden..."
# $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
