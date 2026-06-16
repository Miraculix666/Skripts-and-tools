<#
.SYNOPSIS
    Universal-MasterUpdate.ps1 - Enterprise Update Suite (Zero-Bloatware Edition)

.DESCRIPTION
    Automatisierte Installation von Treibern, Firmware und BIOS-Updates.
    Fokus auf "Clean Installs": Das Skript entpackt ZIP-Archive und versucht
    aktiv, Installer (.exe/.msi) zu entpacken, um nur die reinen .inf-Treiber
    via PnPUtil zu installieren. Das verhindert die Installation von
    unerwuenschter Zusatzsoftware (Bloatware, Tray-Apps).

    FEATURES:
    - ZIP-Support: Entpackt Archive vollautomatisch vor der Analyse.
    - Smart Extraction: Entpackt MSI/EXE (ausser Firmware/WWAN) und archiviert Installer.
    - Safe Driver Cleanup: Loescht blockierende Alt-Treiber.
    - Remote Safe: Ueberspringt Netzwerkkarten-Resets, um PSSessions zu schuetzen.
    - Hardware Fixes: Fujitsu BD-PROCHOT 0,18GHz Fix integriert.
    - OPSI Driver Export: Speichert den sauberen Zustand strukturiert ab.

.PARAMETER SourcePath
    Pfad zu den Update-Dateien. Standard: Ordner des Skripts.

.PARAMETER SkipConfirmation
    Fuer Unattended / SCCM / OPSI Deployments. Unterdrueckt alle Y/N Fragen.

.PARAMETER RemoteMode
    Schuetzt PSSessions vor Verbindungsabbruechen bei der Treiber-Bereinigung.

.VERSION
    7.0 (ZIP Support, Smart Extract, Zero-Bloatware, OPSI-Export)
#>

[CmdletBinding()]
Param(
    [Parameter(Mandatory = $false)]
    [string]$SourcePath = "",

    [Parameter(Mandatory = $false)]
    [switch]$SkipConfirmation,

    [Parameter(Mandatory = $false)]
    [switch]$RemoteMode
)

$ErrorActionPreference = 'Continue'

# =================================================================
# 1. INITIALISIERUNG & PFADE
# =================================================================

$script:RealScriptPath = $MyInvocation.MyCommand.Definition
if ([string]::IsNullOrWhiteSpace($script:RealScriptPath)) { $script:RealScriptPath = $PSCommandPath }

if ([string]::IsNullOrWhiteSpace($SourcePath)) {
    $SourcePath = Split-Path -Parent $script:RealScriptPath
    if ([string]::IsNullOrWhiteSpace($SourcePath)) { $SourcePath = (Get-Location).Path }
}

# Verzeichnisse anlegen
$LogDir = "C:\Daten\Log\UniversalUpdater"
$ArchiveDir = Join-Path $SourcePath "_Archived_Installers"
$ExtractedDir = Join-Path $SourcePath "_Extracted_Drivers"
$ExportDir = "C:\Daten\Drivers\Exported"

foreach ($dir in @($LogDir, $ArchiveDir, $ExtractedDir, $ExportDir)) {
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
}

$LogFile = Join-Path $LogDir "UniversalMasterUpdate_Log.txt"
$PnpExe = "$env:windir\System32\pnputil.exe"
if (-not (Test-Path $PnpExe)) { $PnpExe = "$env:windir\sysnative\pnputil.exe" }

$ForceExeArgs = @("/S", "/Q", "/Silent", "/Quiet", "/Force", "/B", "/ForceIt", "/norestart")
$ForceMsiArgs = @("/qn", "/norestart", "REBOOT=ReallySuppress")

Function Write-Log ([string]$Message, [string]$Level = "INFO") {
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $Color = switch ($Level) {
        "DEBUG" { "DarkGray" } "WARN" { "Yellow" } "ERROR" { "Red" }
        "OK" { "Green" } "EXEC" { "Magenta" } "PLAN" { "Cyan" } Default { "White" }
    }
    $LogLine = "[$Timestamp] [$Level] $Message"
    try { Add-Content -Path $LogFile -Value $LogLine -ErrorAction SilentlyContinue } catch {}
    Write-Host $LogLine -ForegroundColor $Color
}

Write-Log "=====================================================================" "INFO"
Write-Log " UNIVERSAL MASTER UPDATE SUITE v7.0 (Zero-Bloatware)" "INFO"
Write-Log " Host: $env:COMPUTERNAME | Remote Mode: $RemoteMode" "INFO"
Write-Log " Quelle: $SourcePath" "INFO"
Write-Log "=====================================================================" "INFO"

# =================================================================
# 2. ENTPACK-ROUTINEN (ZIP & SMART EXTRACT)
# =================================================================

Function Invoke-ArchiveExtraction {
    Write-Log "Suche nach ZIP-Archiven..." "EXEC"
    $ZipFiles = Get-ChildItem -Path $SourcePath -Recurse -Filter "*.zip" | Where-Object { $_.FullName -notmatch "_Archived_Installers" }

    foreach ($zip in $ZipFiles) {
        $TargetFolder = Join-Path $ExtractedDir ($zip.BaseName -replace '[^a-zA-Z0-9]', '_')
        Write-Log "Entpacke Archiv: $($zip.Name) -> $TargetFolder" "INFO"
        try {
            Expand-Archive -Path $zip.FullName -DestinationPath $TargetFolder -Force -ErrorAction Stop
            Write-Log "Verschiebe $($zip.Name) ins Archiv." "DEBUG"
            Move-Item -Path $zip.FullName -Destination $ArchiveDir -Force -ErrorAction SilentlyContinue
        }
        catch { Write-Log "Fehler beim Entpacken von $($zip.Name): $($_.Exception.Message)" "ERROR" }
    }
}

Function Invoke-SmartExtraction {
    Write-Log "Suche nach entpackbaren Installern (.exe / .msi)..." "EXEC"

    $Installers = Get-ChildItem -Path $SourcePath -Recurse -Include "*.exe", "*.msi" | Where-Object {
        $_.FullName -notmatch "_Archived_Installers" -and $_.Name -notmatch "uninstall|tools|setup_admin"
    }

    foreach ($app in $Installers) {
        $nameLower = $app.Name.ToLower()

        # Firmware, BIOS und WWAN (Brauchen Dienste) duerfen NICHT entpackt werden!
        if ($nameLower -match "bios|uefi|firmware|retimer|ec|thunderbolt_fw|flash|wwan|sierra|fibocom") {
            Write-Log "Ueberspringe Extraktion fuer systemkritische Komponente: $($app.Name)" "DEBUG"
            continue
        }

        $TargetFolder = Join-Path $ExtractedDir ($app.BaseName -replace '[^a-zA-Z0-9]', '_')
        if (-not (Test-Path $TargetFolder)) { New-Item -ItemType Directory -Path $TargetFolder -Force | Out-Null }

        $ExtractSuccess = $false
        Write-Log "Versuche Silent-Extraktion von: $($app.Name)" "INFO"
        try {
            if ($app.Extension -eq ".msi") {
                $Proc = Start-Process -FilePath "msiexec.exe" -ArgumentList "/a `"$($app.FullName)`" /qn TARGETDIR=`"$TargetFolder`"" -Wait -PassThru -WindowStyle Hidden
                if ($Proc.ExitCode -eq 0) { $ExtractSuccess = $true }
            }
            elseif ($app.Extension -eq ".exe") {
                $ExtractArgs = @("/extract=`"$TargetFolder`"", "-extract `"$TargetFolder`"", "/s /e /f `"$TargetFolder`"")
                foreach ($arg in $ExtractArgs) {
                    if ($ExtractSuccess) { break }
                    $Proc = Start-Process -FilePath $app.FullName -ArgumentList $arg -Wait -PassThru -WindowStyle Hidden
                    if ($Proc.ExitCode -eq 0 -and (Get-ChildItem $TargetFolder).Count -gt 0) { $ExtractSuccess = $true }
                }
            }

            # Pruefen ob tatsaechlich .inf Dateien entpackt wurden
            $InfFound = Get-ChildItem -Path $TargetFolder -Recurse -Filter "*.inf" -ErrorAction SilentlyContinue

            if ($ExtractSuccess -and $InfFound) {
                Write-Log "Extraktion ERFOLGREICH. $($InfFound.Count) reine Treiber (.inf) gefunden." "OK"
                Write-Log "Verschiebe originalen Installer ins Archiv (wird NICHT ausgefuehrt)." "WARN"
                Move-Item -Path $app.FullName -Destination $ArchiveDir -Force -ErrorAction SilentlyContinue
            }
            else {
                Write-Log "Extraktion fehlgeschlagen oder keine Treiber enthalten. Wird regulaer silent installiert." "DEBUG"
                Remove-Item -Path $TargetFolder -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
        catch { Write-Log "Fehler bei der Extraktionsroutine fuer $($app.Name)" "ERROR" }
    }
}

# =================================================================
# 3. PRE-FLIGHT CHECK
# =================================================================

# Dateiblockaden entfernen VOR der Analyse
Get-ChildItem -Path $SourcePath -Recurse -File -ErrorAction SilentlyContinue | Unblock-File -ErrorAction SilentlyContinue

# VOR dem Check die Extraktion laufen lassen
Invoke-ArchiveExtraction
Invoke-SmartExtraction

$Plan = @{ HasDrivers = $false; FirmwareApps = @(); BiosApps = @() }
Write-Log "`nErstelle Execution-Plan aus verbleibenden aktiven Dateien..." "INFO"

if (Test-Path $SourcePath) {
    if (Get-ChildItem -Path $SourcePath -Recurse -Filter "*.inf" -ErrorAction SilentlyContinue | Select-Object -First 1) {
        $Plan.HasDrivers = $true
    }

    $AllApps = Get-ChildItem -Path $SourcePath -Recurse -Include "*.exe", "*.msi" | Where-Object {
        $_.FullName -notmatch "_Archived_Installers" -and $_.Name -notmatch "uninstall|tools|setup_admin"
    }

    foreach ($app in $AllApps) {
        if ($app.FullName.ToLower() -match "bios|uefi") { $Plan.BiosApps += $app }
        else { $Plan.FirmwareApps += $app }
    }
}

Write-Log "`n--- EXECUTION PLAN ---" "PLAN"
Write-Log "PnP-Treiber (.inf) verfuegbar: $($Plan.HasDrivers) (Inkl. entpackter ZIPs/EXEs)" "PLAN"
Write-Log "Zwingend auszufuehrende Firmware/WWAN-Installer: $($Plan.FirmwareApps.Count)" "PLAN"
foreach ($f in $Plan.FirmwareApps) { Write-Log " -> $($f.Name)" "DEBUG" }
Write-Log "Auszufuehrende BIOS-Pakete: $($Plan.BiosApps.Count)" "PLAN"
Write-Log "------------------------`n" "PLAN"

if (-not $SkipConfirmation) {
    $Ans = Read-Host "Moechten Sie diesen Ausfuehrungsplan jetzt anwenden? (J/N)"
    if ($Ans -notmatch "^[JjYy]") { Write-Log "Abbruch durch Benutzer." "WARN"; Exit }
}

# =================================================================
# 4. KERNFUNKTIONEN (INSTALLATION & FIXES)
# =================================================================

Function Invoke-SafeDriverCleanup {
    Write-Log "Starte Bereinigung blockierender OEM-Treiber..." "EXEC"

    $TargetKeywords = "Thunderbolt|UCSI|Retimer|Fibocom|Sierra|Realtek"
    $TargetMakers = "Intel"

    try {
        $BadDrivers = Get-CimInstance Win32_PnPSignedDriver -ErrorAction Stop | Where-Object {
            ($_.DeviceName -match $TargetKeywords) -or ($_.Manufacturer -match $TargetMakers -and $_.DeviceName -match "USB|Type-C|System|Audio|WWAN")
        }

        $Deleted = 0
        foreach ($Driver in $BadDrivers) {
            if ($Driver.InfName -match "^oem\d+\.inf$") {
                if ($RemoteMode -and $Driver.DeviceName -match "Wireless|Wi-Fi|Ethernet|GbE|WWAN") {
                    Write-Log "RemoteMode-Schutz: Ueberspringe Netzwerkkarte $($Driver.InfName)" "WARN"
                    continue
                }
                Write-Log "Loesche: $($Driver.InfName) ($($Driver.DeviceName))" "DEBUG"
                $Proc = Start-Process -FilePath $PnpExe -ArgumentList @("/delete-driver", $Driver.InfName, "/uninstall", "/force") -PassThru -Wait -WindowStyle Hidden
                if ($Proc.ExitCode -eq 0 -or $Proc.ExitCode -eq 3010) { $Deleted++ }
            }
        }
        Write-Log "Cleanup beendet. $Deleted Treiber hart geloescht." "OK"
    }
    catch { Write-Log "Fehler beim WMI-Zugriff (Cleanup)." "ERROR" }
}

Function Invoke-DriverInstallation {
    if (-not $Plan.HasDrivers) { return }
    Write-Log "Injiziere reine Treiberdateien (.inf) ins System..." "EXEC"

    $InfSearchPath = Join-Path -Path $SourcePath -ChildPath "*.inf"
    try {
        $Proc = Start-Process -FilePath $PnpExe -ArgumentList @("/add-driver", $InfSearchPath, "/subdirs", "/install") -PassThru -Wait -WindowStyle Hidden
        Write-Log "Treiber-Installation beendet. ExitCode: $($Proc.ExitCode)" "OK"
    }
    catch { Write-Log "Fehler bei PnPUtil." "ERROR" }
}

Function Invoke-HardwareFixes {
    Write-Log "Wende Hardware-Fixes an..." "EXEC"
    $Sys = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue

    # Fujitsu 0,18 GHz CPU Throttling Fix (BD PROCHOT)
    if ($Sys -and $Sys.Manufacturer -match "Fujitsu") {
        Write-Log "Fujitsu Anti-Throttling Profil wird angewendet..." "INFO"
        try {
            $ActiveScheme = (powercfg /getactivescheme) -match 'GUID: (.*?) \('; $G = $Matches[1]
            if ($G) {
                # Min & Max CPU 100%, Active Cooling
                powercfg /setacvalueindex $G 54533251-82be-4824-96c1-47b60b740d00 893dee8e-2bef-41e0-89c6-b55d0929964c 100
                powercfg /setdcvalueindex $G 54533251-82be-4824-96c1-47b60b740d00 893dee8e-2bef-41e0-89c6-b55d0929964c 100
                powercfg /setacvalueindex $G 54533251-82be-4824-96c1-47b60b740d00 bc5038f7-23e0-4960-96da-33abaf5935ec 100
                powercfg /setdcvalueindex $G 54533251-82be-4824-96c1-47b60b740d00 bc5038f7-23e0-4960-96da-33abaf5935ec 100
                powercfg /setacvalueindex $G 54533251-82be-4824-96c1-47b60b740d00 94d3a615-a899-4ac5-ae2b-e4d8f634360f 1
                powercfg /setactive $G
                Write-Log "Fujitsu Power-Fix aktiv." "OK"
            }
        }
        catch {}
    }

    # UCSI Reset
    try {
        $UCSI = Get-PnpDevice -FriendlyName "*UCSI*" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($UCSI -and -not $RemoteMode) {
            Write-Log "Setze UCSI USB-C Schnittstelle zurueck..." "INFO"
            $UCSI | Disable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
            $UCSI | Enable-PnpDevice -Confirm:$false -ErrorAction SilentlyContinue
        }
    }
    catch {}
}

Function Invoke-FirmwareAndBios {
    if ($Plan.FirmwareApps.Count -gt 0) {
        Write-Log "Starte verbleibende Firmware/WWAN-Updates (Unattended)..." "EXEC"
        foreach ($app in $Plan.FirmwareApps) {
            Write-Log " -> Fuehre aus: $($app.Name)" "INFO"
            try {
                if ($app.Extension -eq ".msi") {
                    $ArgsToUse = @("/i", "`"$($app.FullName)`"") + $ForceMsiArgs
                    $Proc = Start-Process -FilePath "msiexec.exe" -ArgumentList $ArgsToUse -Wait -PassThru -WindowStyle Hidden
                }
                else {
                    $Proc = Start-Process -FilePath $app.FullName -ArgumentList $ForceExeArgs -Wait -PassThru -WindowStyle Hidden
                }
            }
            catch { Write-Log "    Fehler: $($_.Exception.Message)" "ERROR" }
        }
    }

    if ($Plan.BiosApps.Count -gt 0) {
        Write-Log "Starte BIOS-Update..." "EXEC"
        $bios = $Plan.BiosApps[0]
        Write-Log " -> Fuehre aus: $($bios.Name)" "WARN"
        try {
            $Proc = Start-Process -FilePath $bios.FullName -ArgumentList $ForceExeArgs -Wait -PassThru -WindowStyle Hidden
        }
        catch { Write-Log "    Fehler: $($_.Exception.Message)" "ERROR" }
    }
}

Function Export-CleanDriverStore {
    Write-Log "Exportiere den sauberen Systemzustand fuer OPSI/SCCM..." "EXEC"
    $SysInfo = Get-CimInstance Win32_ComputerSystem
    $Mfg = $SysInfo.Manufacturer -replace '[\\/:*?"<>|]', '_'
    $Mdl = $SysInfo.Model -replace '[\\/:*?"<>|]', '_'

    $FinalExportPath = Join-Path $ExportDir (Join-Path $Mfg $Mdl)
    if (-not (Test-Path $FinalExportPath)) { New-Item -ItemType Directory -Path $FinalExportPath -Force | Out-Null }

    try {
        Write-Log "Exportiere Treiber nach: $FinalExportPath" "INFO"
        Export-WindowsDriver -Online -Destination $FinalExportPath -ErrorAction Stop | Out-Null
        $exported = (Get-ChildItem -Path $FinalExportPath).Count
        Write-Log "ERFOLG: $exported Treiberpakete exportiert." "OK"
    }
    catch { Write-Log "Fehler beim Treiber-Export: $($_.Exception.Message)" "ERROR" }
}

# =================================================================
# 5. ABLAUF-STEUERUNG
# =================================================================

Invoke-SafeDriverCleanup
Invoke-DriverInstallation
Invoke-HardwareFixes
Invoke-FirmwareAndBios

if (-not $SkipConfirmation) {
    $ExportAns = Read-Host "`nMoechten Sie die aktuell aktiven Treiber fuer die OPSI-Integration (byAudit) exportieren? (J/N)"
    if ($ExportAns -match "^[JjYy]") { Export-CleanDriverStore }
}
elseif ($SkipConfirmation) {
    # Bei Silent-Deployments speichern wir den Clean-State automatisch
    Export-CleanDriverStore
}

Write-Log "=====================================================================" "OK"
Write-Log " UNIVERSAL UPDATE SUITE ERFOLGREICH ABGESCHLOSSEN." "OK"
Write-Log " Log: $LogFile" "INFO"
Write-Log "=====================================================================" "OK"

if (-not $SkipConfirmation -and -not $RemoteMode) {
    $Ans = Read-Host "Moechten Sie das System jetzt neu starten? (J/N)"
    if ($Ans -match "^[JjYy]") { Restart-Computer -Force }
}
