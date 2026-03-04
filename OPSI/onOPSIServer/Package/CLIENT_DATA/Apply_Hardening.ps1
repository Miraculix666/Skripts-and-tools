# Filename: PS_Apply_Harden_Policies.ps1
# Description: Windows 11 Master-Hardening, Debloating & UI-Optimization Script V9.0 (Ultimate Edition)
# Compatibility: PowerShell 5.1+, Windows 10/11, Windows Server 2019/2022
# Usage: Standalone, OPSI, Unattended.xml

<#
.SYNOPSIS
    Umfassendes Hardening-Skript für Windows 11 Umgebungen.
.DESCRIPTION
    Dieses Skript führt Sicherheits-Härtungen gemäß BSI-Empfehlungen durch, entfernt unerwünschte AppX-Pakete
    dauerhaft und passt die Benutzeroberfläche für den professionellen Einsatz an.
    Optimiert für deutsche Enterprise-Umgebungen (Lokalisierung de-DE).
.PARAMETER EnableDefender
    Aktiviert den Defender-Echtzeitschutz und ASR-Regeln im Block-Modus (Standard: Audit/False).
.PARAMETER DriversPathSource
    Pfad zu den zu injizierenden Treibern (Standard: C:\Drivers_Temp).
.PARAMETER SkipDriverInstall
    Überspringt die Treiber-Installation via PnPUtil.
.PARAMETER SilentMode
    Deaktiviert interaktive Bestätigungen (Ideal für OPSI/Unattended).
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)][string]$ModulePrivacy = "True",
    [Parameter(Mandatory = $false)][string]$ModuleSecurity = "True",
    [Parameter(Mandatory = $false)][string]$ModuleDebloat = "True",
    [Parameter(Mandatory = $false)][string]$ModuleUI = "True",
    [Parameter(Mandatory = $false)][string]$EnableDefender = "False",
    [Parameter(Mandatory = $false)][string]$DriversPathSource = "C:\Drivers_Temp",
    [Parameter(Mandatory = $false)][switch]$SkipDriverInstall = $false,
    [Parameter(Mandatory = $false)][switch]$SilentMode = $false
)

# --- Initialisierung & Variablen ---
$ErrorActionPreference = 'Stop'
$VerbosePreference = 'Continue'
$Global:LogFile = Join-Path $env:TEMP "Win11_Hardening_V9_0.log"
$Global:Stats = @{ Success = 0; Warnings = 0; Errors = 0; RebootRequired = $false }

# Lokalisierungseinstellungen für Deutschland
$Culture = [System.Globalization.CultureInfo]::GetCultureInfo("de-DE")
$DateString = Get-Date -Format "dd.MM.yyyy HH:mm:ss"

# Registry-Pfade
$HKLM = "HKLM:\SOFTWARE"
$HKCU = "HKCU:\SOFTWARE"

# --- Hilfsfunktionen ---
function Write-LogEntry {
    param([string]$Message, [string]$Type = "INFO")
    $Timestamp = Get-Date -Format "HH:mm:ss"
    $LogMsg = "[$Timestamp] [$Type] $Message"

    switch ($Type) {
        "ERROR" { Write-Error $Message; $Global:Stats.Errors++ }
        "WARNING" { Write-Warning $Message; $Global:Stats.Warnings++ }
        "SUCCESS" { Write-Host "  [OK] $Message" -ForegroundColor Green; $Global:Stats.Success++ }
        "INFO" { Write-Verbose "INFO: $Message" }
    }

    try { $LogMsg | Out-File -FilePath $Global:LogFile -Append -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
}

function Set-RegValue {
    param($Path, $Name, $Value, $Type = "DWord", $Description = "")
    try {
        if (-not (Test-Path $Path)) { New-Item -Path $Path -Force | Out-Null }
        Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type -Force | Out-Null
        Write-LogEntry "Registry gesetzt: $Description ($Name)" "SUCCESS"
        return $true
    }
    catch {
        Write-LogEntry "Fehler bei Registry $Name : $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# --- Core Module ---

function Invoke-PrivacyHardening {
    Write-LogEntry "Starte Modul: Privacy & Telemetrie (BSI-konform)..."

    $Policies = @(
        @{ Path = "$HKLM\Policies\Microsoft\Windows\DataCollection"; Name = 'AllowTelemetry'; Value = 0; Desc = 'Telemetrie deaktiviert' },
        @{ Path = "$HKLM\Policies\Microsoft\Windows\AdvertisingInfo"; Name = 'DisabledByGroupPolicy'; Value = 1; Desc = 'Werbe-ID deaktiviert' },
        @{ Path = "$HKLM\Policies\Microsoft\Windows\DeliveryOptimization"; Name = 'DODownloadMode'; Value = 99; Desc = 'Übermittlungsoptimierung aus (99)' },
        @{ Path = "$HKLM\Policies\Microsoft\MRT"; Name = 'DontReportInfectionInformation'; Value = 1; Desc = 'MRT Telemetrie aus' },
        @{ Path = "$HKLM\Policies\Microsoft\Windows\WcmSvc\Local"; Name = 'AllowWiFiHotSpotReporting'; Value = 0; Desc = 'WiFi Sense deaktiviert' },
        @{ Path = "$HKLM\Policies\Microsoft\Windows\Windows Search"; Name = 'AllowCortana'; Value = 0; Desc = 'Cortana Suche deaktiviert' },
        @{ Path = "$HKLM\Policies\Microsoft\Windows\Windows Search"; Name = 'ConnectedSearchUseWeb'; Value = 0; Desc = 'Websuche in Startmenü aus' }
    )

    foreach ($P in $Policies) { Set-RegValue -Path $P.Path -Name $P.Name -Value $P.Value -Desc $P.Desc }
}

function Invoke-SecurityHardening {
    Write-LogEntry "Starte Modul: Erweiterte Sicherheit (LSA, CG, ASR)..."

    # LSA Protection
    Set-RegValue -Path "$HKLM\SYSTEM\CurrentControlSet\Control\Lsa" -Name 'RunAsPPL' -Value 1 -Desc 'LSA Protection aktiv'

    # Credential Guard (Prüfung Hardware-Voraussetzung)
    try {
        $dg = Get-CimInstance -Namespace root\Microsoft\Windows\DeviceGuard -ClassName MSFT_DeviceGuard
        if ($dg.VirtualizationBasedSecurityStatus -ge 1) {
            Set-RegValue -Path "$HKLM\SYSTEM\CurrentControlSet\Control\Lsa" -Name 'LsaCfgFlags' -Value 1 -Desc 'Credential Guard Hardware-validiert'
        }
        else {
            Write-LogEntry "Credential Guard nicht möglich (VBS nicht aktiv)." "WARNING"
        }
    }
    catch {
        Write-LogEntry "Hardware-Check für Credential Guard fehlgeschlagen." "WARNING"
    }

    # ASR Regeln
    $ASRMode = if ($EnableDefender -eq "True") { "1" } else { "2" } # 1=Block, 2=Audit
    $ASRPath = "$HKLM\Policies\Microsoft\Windows Defender\Windows Defender Exploit Guard\ASR\Rules"
    $Rules = @("be9ba2d9-53ea-4cdc-84e5-9b1eeee46550", "d4f940ab-401b-4efc-aadc-ad5f3c50688a", "9e6c4e1f-7d60-472f-ba1a-a39ef669e4b2")
    foreach ($R in $Rules) { Set-RegValue -Path $ASRPath -Name $R -Value $ASRMode -Type "String" -Desc "ASR Regel $R" }
}

function Invoke-DauerhaftesDebloating {
    Write-LogEntry "Starte Modul: Permanentes Debloating (AppX & Provisioned)..."

    # Liste der zu entfernenden Apps
    $AppList = @(
        "*Teams*", "*Xbox*", "*Copilot*", "*Clipchamp*", "*OneDrive*",
        "*OneNote*", "*News*", "*Weather*", "*Zune*", "*Office.OneNote*",
        "*Bing*", "*Solitaire*", "*People*", "*CommunicationApps*"
    )

    # 1. Provisioned Packages (verhindert Neuinstallation bei neuen Usern)
    $Provisioned = Get-AppxProvisionedPackage -Online
    foreach ($AppPattern in $AppList) {
        $Matches = $Provisioned | Where-Object { $_.DisplayName -like $AppPattern }
        foreach ($Match in $Matches) {
            Write-LogEntry "Entferne Provisioned App: $($Match.DisplayName)" "INFO"
            Remove-AppxProvisionedPackage -Online -PackageName $Match.PackageName -ErrorAction SilentlyContinue | Out-Null
        }
    }

    # 2. Aktuelle Benutzer-Pakete
    foreach ($AppPattern in $AppList) {
        Get-AppxPackage -AllUsers | Where-Object { $_.Name -like $AppPattern } | ForEach-Object {
            Write-LogEntry "Entferne Benutzer-App: $($_.Name)" "INFO"
            Remove-AppxPackage -Package $_.PackageFullName -AllUsers -ErrorAction SilentlyContinue | Out-Null
        }
    }

    # 3. Store Auto-Download Sperre (um Nachladen zu verhindern)
    Set-RegValue -Path "$HKLM\Software\Policies\Microsoft\WindowsStore" -Name "AutoDownload" -Value 2 -Desc "Store Auto-Update deaktiviert"
}

function Invoke-UIAndSystem {
    Write-LogEntry "Starte Modul: UI-Anpassungen & System..."

    # Windows 10 Context Menu (Classic)
    Set-RegValue -Path "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32" -Name "" -Value "" -Type "String" -Desc "Classic Context Menu"

    # Explorer Einstellungen
    $Exp = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    Set-RegValue -Path $Exp -Name "Start_TaskbarAlignment" -Value 0 -Desc "Taskbar linksbündig"
    Set-RegValue -Path $Exp -Name "Hidden" -Value 1 -Desc "Versteckte Dateien anzeigen"
    Set-RegValue -Path $Exp -Name "HideFileExt" -Value 0 -Desc "Dateiendungen anzeigen"

    # Power Settings (OPSI Standard)
    Set-RegValue -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power" -Name 'HiberbootEnabled' -Value 0 -Desc "Fastboot aus"

    # Treiber Injektion
    if (-not $SkipDriverInstall -and (Test-Path $DriversPathSource)) {
        Write-LogEntry "Injektiere Treiber aus $DriversPathSource..." "INFO"
        try {
            pnputil /add-driver "$DriversPathSource\*.inf" /install /subdirs | Out-Null
            Write-LogEntry "Treiber-Injektion erfolgreich." "SUCCESS"
        }
        catch {
            Write-LogEntry "Fehler bei Treiber-Injektion: $($_.Exception.Message)" "WARNING"
        }
    }
}

# --- Hauptausführung ---
function Main {
    Write-Host "`n--- PS-Coding Hardening Suite V9.0 ---" -ForegroundColor Cyan
    Write-LogEntry "Skriptstart auf $(hostname) durch $(whoami)"

    if (-not $SilentMode) {
        Write-Host "VORSICHT: Dieses Skript ändert tiefgreifende Systemeinstellungen." -ForegroundColor Yellow
        $Confirm = Read-Host "Fortfahren? (J/N)"
        if ($Confirm -ne "J") { Write-LogEntry "Abbruch durch Benutzer." "WARNING"; return }
    }

    if ($ModulePrivacy -eq "True") { Invoke-PrivacyHardening }
    if ($ModuleSecurity -eq "True") { Invoke-SecurityHardening }
    if ($ModuleDebloat -eq "True") { Invoke-DauerhaftesDebloating }
    if ($ModuleUI -eq "True") { Invoke-UIAndSystem }

    Write-Host "`n--- Zusammenfassung ---" -ForegroundColor Cyan
    Write-Host "Erfolgreich: $($Global:Stats.Success)" -ForegroundColor Green
    Write-Host "Warnungen:   $($Global:Stats.Warnings)" -ForegroundColor Yellow
    Write-Host "Fehler:      $($Global:Stats.Errors)" -ForegroundColor Red
    Write-Host "Logdatei:    $Global:LogFile"

    if ($Global:Stats.Errors -eq 0) {
        exit 0
    }
    else {
        exit 1
    }
}

# Startet die Main-Funktion
Main
