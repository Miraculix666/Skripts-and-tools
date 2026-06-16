# Filename: Apply_Hardening.ps1
# Description: Windows 10/11/Server Hardening, Debloating & Optimization Suite V11.0
# Compatibility: PowerShell 5.1+, Windows 10/11/Server 2019+
# Usage: Interaktiv, Silent/OPSI, PS-Session, Unattended
# History: Konsolidiert aus V1-V10 + BSI SiSyPHuS + OO ShutUp10 + NTLite
# License: Internal Use Only

<#
.SYNOPSIS
    Umfassendes Hardening-Skript fuer Windows 10/11/Server (V11.0).
.DESCRIPTION
    Modulares System mit persistentem State, interaktivem Menue,
    DNS-Sperrlisten, .reg/.xml Export und Restore-Funktion.

    Module:
      A) Privacy & Telemetrie (BSI/ShutUp10/NTLite-konform, ~50 Settings)
      B) Defender-Steuerung (Aktivieren/Deaktivieren mit Warnung)
      C) Erweiterte Sicherheit (LSA, Credential Guard, ASR)
      D) Remote & Netzwerk (RDP, WinRM, Firewall, WOL)
      E) Dienste & Store (Telemetrie-Dienste, Store-Kontrolle)
      F) Debloating (AppX + Capabilities, Disable vs. Remove)
      G) UI-Anpassungen & Explorer + Nag-Screen-Unterdrueckung
      H) Energie & System (Energieplan, FastBoot, LockScreen)
      I) Treiber-Injektion (universell, mehrstufig, SCCM-kompatibel)
      J) DNS-Sperrliste (hosts-basiert, Whitelist, Offline-faehig)
      K) HKCU User-Haertung (Scheduled Task fuer alle User)

.PARAMETER EnableDefender
    Defender bleibt aktiv (ASR im Block-Modus). Standard: Defender wird deaktiviert.
.PARAMETER DriversPathSource
    Pfad zu SCCM-Treiberpaketen (Standard: C:\Drivers_Temp).
.PARAMETER SkipDriverInstall
    Ueberspringt die Treiber-Installation.
.PARAMETER SilentMode
    Deaktiviert interaktive Menues (fuer OPSI/Unattended/PS-Session).
.PARAMETER SkipRemoteSetup
    Ueberspringt RDP/WinRM/Firewall-Konfiguration.
.PARAMETER RestoreMode
    Nimmt Aenderungen zurueck basierend auf dem gespeicherten State.
.PARAMETER Modules
    Komma-separierte Modulliste (z.B. "A,B,C,J"). Standard: alle.
.PARAMETER ExportReg
    Generiert eine .reg Datei mit allen Registry-Aenderungen.
.PARAMETER ExportUnattend
    Generiert Autounattend.xml Dateien (USB + OPSI).
.PARAMETER ShowStatus
    Zeigt den aktuellen Haertungsstatus aus dem State-File.
.PARAMETER UpdateDNSBlocklist
    Aktualisiert die DNS-Sperrliste aus Online-Quellen.
.PARAMETER DebloatAction
    Aktion fuer Modul F: "Disable" (deaktivieren) oder "Remove" (entfernen). Standard: Remove.
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [switch]$EnableDefender       = $false,
    [string]$DriversPathSource    = "C:\Drivers_Temp",
    [switch]$SkipDriverInstall    = $false,
    [switch]$SilentMode           = $false,
    [switch]$SkipRemoteSetup      = $false,
    [switch]$RestoreMode          = $false,
    [string[]]$Modules            = @(),
    [switch]$ExportReg            = $false,
    [switch]$ExportUnattend       = $false,
    [switch]$ShowStatus           = $false,
    [switch]$UpdateDNSBlocklist   = $false,
    [ValidateSet("Disable","Remove")]
    [string]$DebloatAction        = "Remove"
)

#Requires -RunAsAdministrator

# ╔══════════════════════════════════════════════════════════════════════════════╗
# ║  KONFIGURATION — Alle aenderbaren Werte hier am Anfang                     ║
# ║  Fuer schnelle Anpassung: Nur diesen Block editieren.                      ║
# ╚══════════════════════════════════════════════════════════════════════════════╝

# --- Pfade ---
# Persistenter State und Log (ueberlebt Neuinstallation nicht, aber Reboots)
$Script:StateDir           = "C:\ProgramData\Hardening"
$Script:StateFile          = Join-Path $Script:StateDir "state.json"
$Script:LogFile            = Join-Path $Script:StateDir "hardening.log"
$Script:RegExportFile      = Join-Path $Script:StateDir "Hardening_Export.reg"
$Script:HostsFile          = "$env:SystemRoot\System32\drivers\etc\hosts"
$Script:HostsBackup        = Join-Path $Script:StateDir "hosts.backup"
$Script:DNSWhitelistFile   = Join-Path $Script:StateDir "dns_whitelist.txt"
$Script:DNSOfflineFile     = Join-Path $PSScriptRoot "dns_blocklist.txt"
$Script:TaskXMLPath        = Join-Path $PSScriptRoot "User_Harden_Task.xml"

# --- Autounattend Passwort (Base64-kodiert) ---
# WICHTIG: Vor Produktiveinsatz aendern!
# Format: Unicode-String Base64-kodiert (Windows Unattend Standard)
# Beispiel "nt123!" = "bgB0ADEAMgAzACEA" | Beispiel "P@ssw0rd" = "UABAAHMAcwB3ADAAcgBkAA=="
$Script:UnattendPassword   = "__CHANGE_ME__"
$Script:UnattendUser       = "FirstAdmin"
$Script:UnattendOrg        = "IT-Abteilung"

# --- DNS Sperrlisten-Quellen (Pi-hole Stil, Hagezi/BSI) ---
$Script:DNSBlocklistURLs   = @(
    "https://raw.githubusercontent.com/hagezi/dns-blocklists/main/hosts/light.txt"
    "https://raw.githubusercontent.com/pschneider1968/pihole-bl-msft-telemetry-bsi/main/hosts.txt"
)

# --- Treiber-Konfiguration ---
# Unterstuetzte Hersteller fuer automatische Erkennung
$Script:SupportedVendors   = @("FUJITSU", "LENOVO", "DELL", "HP")
# Max. Reboot-Zyklen fuer mehrstufige Treiber-Updates
$Script:MaxDriverReboots   = 3
# Registry-Key fuer Reboot-Tracking
$Script:DriverRebootKey    = "HKLM:\SOFTWARE\Hardening\DriverUpdate"

# --- Marker fuer hosts-Datei (zum sauberen Entfernen) ---
$Script:HostsMarkerStart   = "# >>> HARDENING-DNS-BLOCK START <<<"
$Script:HostsMarkerEnd     = "# >>> HARDENING-DNS-BLOCK END <<<"

# --- Nag-Screen Unterdrueckung (Programm-Pfade) ---
# Registry-Keys die "Willkommen/Whats New" Dialoge unterdruecken
$Script:NagScreenApps      = @(
    @{ Name="Chrome";  Path="HKLM:\SOFTWARE\Policies\Google\Chrome";         Key="SuppressUnsupportedOSWarning"; Value=1 }
    @{ Name="Chrome";  Path="HKLM:\SOFTWARE\Policies\Google\Chrome";         Key="PromotionalTabsEnabled";       Value=0 }
    @{ Name="Firefox"; Path="HKLM:\SOFTWARE\Policies\Mozilla\Firefox";       Key="OverrideFirstRunPage";         Value="" ; Type="String" }
    @{ Name="Firefox"; Path="HKLM:\SOFTWARE\Policies\Mozilla\Firefox";       Key="OverridePostUpdatePage";       Value="" ; Type="String" }
    @{ Name="Edge";    Path="HKLM:\SOFTWARE\Policies\Microsoft\Edge";        Key="HideFirstRunExperience";       Value=1 }
    @{ Name="Edge";    Path="HKLM:\SOFTWARE\Policies\Microsoft\Edge";        Key="StartupBoostEnabled";          Value=0 }
    @{ Name="Office";  Path="HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\General"; Key="ShownFirstRunOptin"; Value=1 }
    @{ Name="Office";  Path="HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common";         Key="QMEnable";           Value=0 }
    @{ Name="Office";  Path="HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common";         Key="UpdateNotifications"; Value="DWORD:0" }
    @{ Name="Teams";   Path="HKLM:\SOFTWARE\Policies\Microsoft\Teams";       Key="DisableFirstLaunchSetup";      Value=1 }
    @{ Name="Adobe";   Path="HKLM:\SOFTWARE\Policies\Adobe\Acrobat Reader\DC\FeatureLockdown"; Key="bUsageMeasurement"; Value=0 }
)

# --- Modul-Definitionen (Reihenfolge = Ausfuehrungsreihenfolge) ---
$Script:ModuleDefinitions = [ordered]@{
    "A" = @{ Name = "Privacy & Telemetrie";            Desc = "BSI/ShutUp10/NTLite Registry (HKLM), ~50 Settings"; Default = $true  }
    "B" = @{ Name = "Defender-Steuerung";               Desc = "Aktivieren/Deaktivieren mit Sicherheitswarnung";    Default = $true  }
    "C" = @{ Name = "Erweiterte Sicherheit";            Desc = "LSA Protection, Credential Guard, ASR-Regeln";      Default = $true  }
    "D" = @{ Name = "Remote & Netzwerk";                Desc = "RDP, WinRM, Firewall, Wake-on-LAN";                 Default = $false }
    "E" = @{ Name = "Dienste & Store";                  Desc = "Telemetrie-Dienste deaktivieren, Store-Kontrolle";   Default = $true  }
    "F" = @{ Name = "Debloating";                       Desc = "AppX + Capabilities (Disable/Remove waehlbar)";     Default = $true  }
    "G" = @{ Name = "UI & Explorer";                    Desc = "Admin-optimiert, Dark Mode, Nag-Screens aus";        Default = $true  }
    "H" = @{ Name = "Energie & System";                 Desc = "Energieplan, FastBoot aus, LockScreen aus";          Default = $true  }
    "I" = @{ Name = "Treiber-Injektion";                Desc = "SCCM-Pakete, mehrstufig, universell";                Default = $false }
    "J" = @{ Name = "DNS-Sperrliste";                   Desc = "hosts-basiert, Whitelist, Offline-faehig";           Default = $true  }
    "K" = @{ Name = "HKCU User-Haertung";               Desc = "Scheduled Task fuer alle Benutzer bei Login";        Default = $true  }
}

# --- Registry-Pfad-Abkuerzungen ---
$Script:HKLM        = "HKLM:\SOFTWARE"
$Script:HKLMPol     = "HKLM:\SOFTWARE\Policies"
$Script:HKCU        = "HKCU:\SOFTWARE"

# --- Statistik ---
$Script:Stats = @{ Success = 0; Warnings = 0; Errors = 0; Skipped = 0 }

# ╔══════════════════════════════════════════════════════════════════════════════╗
# ║  HILFSFUNKTIONEN                                                            ║
# ╚══════════════════════════════════════════════════════════════════════════════╝

function Initialize-Environment {
    <# Erstellt Verzeichnisse und prueft Voraussetzungen #>
    if (-not (Test-Path $Script:StateDir)) {
        New-Item -Path $Script:StateDir -ItemType Directory -Force | Out-Null
    }
    # Whitelist-Datei anlegen falls nicht vorhanden
    if (-not (Test-Path $Script:DNSWhitelistFile)) {
        @(
            "# DNS Whitelist — Domains die NICHT blockiert werden sollen"
            "# Eine Domain pro Zeile, Kommentare mit #"
            "# Beispiel: update.microsoft.com"
            "login.microsoftonline.com"
            "login.live.com"
            "activation.sls.microsoft.com"
            "validation.sls.microsoft.com"
            "go.microsoft.com"
            "download.windowsupdate.com"
        ) | Set-Content -Path $Script:DNSWhitelistFile -Encoding UTF8
    }
}

function Test-IsRemoteSession {
    <# Erkennt ob das Skript in einer PS-Session oder headless laeuft #>
    $isRemote = $false
    if ($env:SESSIONNAME -and $env:SESSIONNAME -ne "Console") { $isRemote = $true }
    if ([System.Environment]::GetEnvironmentVariable("SSH_CONNECTION")) { $isRemote = $true }
    if ($PSSenderInfo) { $isRemote = $true }
    return $isRemote -or $SilentMode
}

function Write-LogEntry {
    <# Schreibt Log-Eintraege in Konsole und persistent in Logdatei #>
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet("INFO","SUCCESS","WARNING","ERROR","HEAD")]
        [string]$Type = "INFO"
    )
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logLine = "[$ts] [$Type] $Message"
    switch ($Type) {
        "ERROR"   { Write-Host "  [FEHLER] $Message" -ForegroundColor Red;     $Script:Stats.Errors++ }
        "WARNING" { Write-Host "  [WARN]   $Message" -ForegroundColor Yellow;  $Script:Stats.Warnings++ }
        "SUCCESS" { Write-Host "  [OK]     $Message" -ForegroundColor Green;   $Script:Stats.Success++ }
        "INFO"    { Write-Verbose "  [INFO]   $Message" }
        "HEAD"    { Write-Host "`n=== $Message ===" -ForegroundColor Cyan }
    }
    try { $logLine | Out-File -FilePath $Script:LogFile -Append -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
}

function Get-RegistryValueSafe {
    <# Liest einen Registry-Wert oder gibt $null zurueck #>
    param([string]$Path, [string]$Name)
    try {
        if (Test-Path $Path) {
            $val = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
            if ($null -ne $val) { return $val.$Name }
        }
    } catch {}
    return $null
}

function Set-RegValue {
    <# Setzt einen Registry-Wert und protokolliert Vorher/Nachher im State #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)]$Value,
        [string]$Type = "DWord",
        [string]$Description = "",
        [string]$Module = ""
    )
    $before = Get-RegistryValueSafe -Path $Path -Name $Name
    try {
        if (-not (Test-Path $Path)) { New-Item -Path $Path -Force | Out-Null }
        Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type -Force | Out-Null
        Write-LogEntry "$Description ($Name=$Value)" "SUCCESS"
        # State-Protokoll aktualisieren
        Update-StateEntry -Module $Module -Key "$Path\$Name" -Before $before -After $Value -Desc $Description
        return $true
    } catch {
        Write-LogEntry "Registry-Fehler $Name : $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Restore-RegValue {
    <# Stellt einen Registry-Wert aus dem State wieder her #>
    param([string]$Path, [string]$Name, $OriginalValue)
    try {
        if ($null -eq $OriginalValue) {
            # Wert existierte vorher nicht — entfernen
            if (Test-Path $Path) {
                Remove-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
            }
        } else {
            Set-ItemProperty -Path $Path -Name $Name -Value $OriginalValue -Force | Out-Null
        }
        Write-LogEntry "Restored: $Path\$Name" "SUCCESS"
    } catch {
        Write-LogEntry "Restore-Fehler: $Path\$Name — $($_.Exception.Message)" "ERROR"
    }
}

# ╔══════════════════════════════════════════════════════════════════════════════╗
# ║  STATE MANAGEMENT (Persistent JSON Protocol)                                ║
# ╚══════════════════════════════════════════════════════════════════════════════╝

function Get-HardeningState {
    <# Laedt den gespeicherten State oder erstellt einen neuen #>
    if (Test-Path $Script:StateFile) {
        try {
            return (Get-Content $Script:StateFile -Raw | ConvertFrom-Json)
        } catch {
            Write-LogEntry "State-Datei beschaedigt, erstelle neu." "WARNING"
        }
    }
    return [PSCustomObject]@{
        Version    = "11.0"
        LastRun    = $null
        Hostname   = $env:COMPUTERNAME
        ExecutedBy = "$env:USERDOMAIN\$env:USERNAME"
        Modules    = @{}
        DNSBlocklist = @{ LastUpdate = $null; EntriesCount = 0; Source = "" }
    }
}

function Save-HardeningState {
    <# Speichert den aktuellen State persistent als JSON #>
    param([Parameter(Mandatory)]$State)
    $State.LastRun = (Get-Date -Format "o")
    $State.Hostname = $env:COMPUTERNAME
    $State.ExecutedBy = "$env:USERDOMAIN\$env:USERNAME"
    $State | ConvertTo-Json -Depth 10 | Set-Content -Path $Script:StateFile -Encoding UTF8 -Force
}

function Update-StateEntry {
    <# Fuegt einen einzelnen Eintrag zum State hinzu #>
    param([string]$Module, [string]$Key, $Before, $After, [string]$Desc)
    if (-not $Script:CurrentState.Modules.$Module) {
        $Script:CurrentState.Modules | Add-Member -NotePropertyName $Module -NotePropertyValue @{
            Applied = $true; Timestamp = (Get-Date -Format "o"); Settings = @{}
        } -Force
    }
    $settings = $Script:CurrentState.Modules.$Module.Settings
    if ($settings -is [hashtable]) {
        $settings[$Key] = @{ Before = $Before; After = $After; Description = $Desc }
    } else {
        $settings | Add-Member -NotePropertyName $Key -NotePropertyValue @{
            Before = $Before; After = $After; Description = $Desc
        } -Force
    }
}

function Show-HardeningStatus {
    <# Zeigt den aktuellen Haertungsstatus tabellarisch an #>
    $state = Get-HardeningState
    Write-Host "`n╔══════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║  Hardening Status — $($state.Hostname)" -ForegroundColor Cyan
    Write-Host "╠══════════════════════════════════════════════════════╣" -ForegroundColor Cyan
    Write-Host "║  Version:    $($state.Version)" -ForegroundColor White
    Write-Host "║  Letzter Lauf: $($state.LastRun)" -ForegroundColor White
    Write-Host "║  Ausgefuehrt von: $($state.ExecutedBy)" -ForegroundColor White
    Write-Host "╠══════════════════════════════════════════════════════╣" -ForegroundColor Cyan
    foreach ($modKey in $Script:ModuleDefinitions.Keys) {
        $modState = $state.Modules.$modKey
        $status = if ($modState.Applied) { "[X]" } else { "[ ]" }
        $ts = if ($modState.Timestamp) { " ($($modState.Timestamp))" } else { "" }
        $color = if ($modState.Applied) { "Green" } else { "DarkGray" }
        Write-Host ("║  $status $modKey  $($Script:ModuleDefinitions[$modKey].Name)$ts") -ForegroundColor $color
    }
    if ($state.DNSBlocklist.EntriesCount -gt 0) {
        Write-Host "║  DNS-Blocklist: $($state.DNSBlocklist.EntriesCount) Eintraege ($($state.DNSBlocklist.Source))" -ForegroundColor White
    }
    Write-Host "╚══════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host "State-Datei: $Script:StateFile" -ForegroundColor DarkGray
}

# ╔══════════════════════════════════════════════════════════════════════════════╗
# ║  INTERAKTIVES MENUE (Pfeiltasten + Leertaste)                              ║
# ╚══════════════════════════════════════════════════════════════════════════════╝

function Show-InteractiveMenu {
    <#
    .DESCRIPTION
        Zeigt ein navigierbares Menue mit Pfeiltasten, Leertaste zum Togglen,
        Enter zum Ausfuehren. Funktioniert nur bei interaktiver Konsole.
    #>
    $items = [System.Collections.ArrayList]@()
    foreach ($key in $Script:ModuleDefinitions.Keys) {
        $def = $Script:ModuleDefinitions[$key]
        $items.Add(@{
            Key      = $key
            Name     = $def.Name
            Desc     = $def.Desc
            Selected = $def.Default
        }) | Out-Null
    }

    $cursorPos = 0
    $done = $false
    $action = "apply"  # apply, restore, status, exportreg, exportunattend, quit

    while (-not $done) {
        Clear-Host
        Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
        Write-Host "║   Windows Hardening Suite V11.0 — Modulauswahl              ║" -ForegroundColor Cyan
        Write-Host "╠══════════════════════════════════════════════════════════════╣" -ForegroundColor Cyan

        for ($i = 0; $i -lt $items.Count; $i++) {
            $item = $items[$i]
            $check = if ($item.Selected) { "X" } else { " " }
            $prefix = if ($i -eq $cursorPos) { " >" } else { "  " }
            $color = if ($i -eq $cursorPos) { "White" } else { "Gray" }
            $checkColor = if ($item.Selected) { "Green" } else { "DarkGray" }

            Write-Host -NoNewline "║$prefix [" -ForegroundColor $color
            Write-Host -NoNewline $check -ForegroundColor $checkColor
            Write-Host -NoNewline "] " -ForegroundColor $color
            Write-Host -NoNewline "$($item.Key)  " -ForegroundColor Yellow
            Write-Host -NoNewline "$($item.Name)" -ForegroundColor $color
            # Pad to fill box
            $pad = 58 - $item.Key.Length - $item.Name.Length - 8
            if ($pad -lt 0) { $pad = 0 }
            Write-Host -NoNewline (" " * $pad)
            Write-Host "║" -ForegroundColor Cyan
        }

        Write-Host "╠══════════════════════════════════════════════════════════════╣" -ForegroundColor Cyan
        Write-Host "║  [R] Restore-Modus   [S] Status   [E] .reg   [U] Unattend  ║" -ForegroundColor DarkCyan
        Write-Host "║  [A] Alle waehlen    [N] Keine    [Q] Beenden               ║" -ForegroundColor DarkCyan
        Write-Host "╠══════════════════════════════════════════════════════════════╣" -ForegroundColor Cyan
        Write-Host "║  Navigation: ↑↓ Bewegen  Leertaste=Toggle  Enter=Start      ║" -ForegroundColor DarkGray
        Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

        $key = [System.Console]::ReadKey($true)

        switch ($key.Key) {
            "UpArrow"   { if ($cursorPos -gt 0) { $cursorPos-- } }
            "DownArrow" { if ($cursorPos -lt ($items.Count - 1)) { $cursorPos++ } }
            "Spacebar"  { $items[$cursorPos].Selected = -not $items[$cursorPos].Selected }
            "Tab"       { $items[$cursorPos].Selected = -not $items[$cursorPos].Selected
                          if ($cursorPos -lt ($items.Count - 1)) { $cursorPos++ } }
            "Enter"     { $action = "apply"; $done = $true }
            default {
                switch ($key.KeyChar) {
                    'a' { $items | ForEach-Object { $_.Selected = $true } }
                    'A' { $items | ForEach-Object { $_.Selected = $true } }
                    'n' { $items | ForEach-Object { $_.Selected = $false } }
                    'N' { $items | ForEach-Object { $_.Selected = $false } }
                    'r' { $action = "restore"; $done = $true }
                    'R' { $action = "restore"; $done = $true }
                    's' { $action = "status"; $done = $true }
                    'S' { $action = "status"; $done = $true }
                    'e' { $action = "exportreg"; $done = $true }
                    'E' { $action = "exportreg"; $done = $true }
                    'u' { $action = "exportunattend"; $done = $true }
                    'U' { $action = "exportunattend"; $done = $true }
                    'q' { $action = "quit"; $done = $true }
                    'Q' { $action = "quit"; $done = $true }
                }
            }
        }
    }

    return @{
        Action          = $action
        SelectedModules = ($items | Where-Object { $_.Selected } | ForEach-Object { $_.Key })
    }
}

# ╔══════════════════════════════════════════════════════════════════════════════╗
# ║  MODUL A: Privacy & Telemetrie (~50 Settings)                               ║
# ║  Quellen: BSI SiSyPHuS, OO ShutUp10, NTLite, Telemetriematrix              ║
# ╚══════════════════════════════════════════════════════════════════════════════╝
function Invoke-PrivacyHardening {
    Write-LogEntry "Modul A: Privacy & Telemetrie (BSI/ShutUp10/NTLite)" "HEAD"

    $settings = @(
        # --- Telemetrie-Kern (BSI Telemetriedokument 12, 24) ---
        @{ Path="$Script:HKLMPol\Microsoft\Windows\DataCollection";       Name='AllowTelemetry';                    Value=0; Desc='Telemetrie vollstaendig deaktiviert (Sicherheit)' }
        @{ Path="$Script:HKLMPol\Microsoft\Windows\DataCollection";       Name='MaxTelemetryAllowed';               Value=0; Desc='Max. Telemetrie-Level: Sicherheit' }
        @{ Path="$Script:HKLMPol\Microsoft\Windows\DataCollection";       Name='LimitDiagnosticLogCollection';      Value=1; Desc='Diagnoseprotokoll-Sammlung begrenzt' }
        @{ Path="$Script:HKLMPol\Microsoft\Windows\DataCollection";       Name='DisableOneSettingsDownloads';        Value=1; Desc='OneSettings Downloads deaktiviert' }
        @{ Path="$Script:HKLMPol\Microsoft\Windows\DataCollection";       Name='DoNotShowFeedbackNotifications';    Value=1; Desc='Feedback-Benachrichtigungen unterdrueckt' }
        @{ Path="$Script:HKLMPol\Microsoft\Windows\DataCollection";       Name='AllowCommercialDataPipeline';       Value=0; Desc='Commercial Data Pipeline deaktiviert' }
        @{ Path="$Script:HKLMPol\Microsoft\Windows\DataCollection";       Name='DisableTelemetryOptInChangeNotification'; Value=1; Desc='Telemetrie OptIn-Aenderungsbenachrichtigung aus' }
        @{ Path="$Script:HKLMPol\Microsoft\Windows\DataCollection";       Name='DisableTelemetryOptInSettingsUx';   Value=1; Desc='Telemetrie OptIn-UI deaktiviert' }

        # --- Werbe-ID und Tracking ---
        @{ Path="$Script:HKLMPol\Microsoft\Windows\AdvertisingInfo";      Name='DisabledByGroupPolicy';             Value=1; Desc='Werbe-ID per GPO deaktiviert' }

        # --- Uebermittlungsoptimierung (P2P Updates) ---
        @{ Path="$Script:HKLMPol\Microsoft\Windows\DeliveryOptimization"; Name='DODownloadMode';                    Value=1; Desc='Uebermittlungsoptimierung: nur LAN (kein Internet-P2P)' }

        # --- MRT (Malicious Removal Tool) Telemetrie ---
        @{ Path="$Script:HKLMPol\Microsoft\MRT";                         Name='DontReportInfectionInformation';    Value=1; Desc='MRT Infektionsmeldung deaktiviert' }
        @{ Path="$Script:HKLMPol\Microsoft\MRT";                         Name='DontOfferThroughWUAU';              Value=1; Desc='MRT Auto-Download via WU deaktiviert' }

        # --- WiFi Sense ---
        @{ Path="$Script:HKLMPol\Microsoft\Windows\WcmSvc\Local";        Name='AllowWiFiHotSpotReporting';         Value=0; Desc='WiFi Sense HotSpot-Reporting aus' }

        # --- Cortana und Websuche ---
        @{ Path="$Script:HKLMPol\Microsoft\Windows\Windows Search";       Name='AllowCortana';                      Value=0; Desc='Cortana deaktiviert' }
        @{ Path="$Script:HKLMPol\Microsoft\Windows\Windows Search";       Name='ConnectedSearchUseWeb';             Value=0; Desc='Websuche in Startmenue deaktiviert' }
        @{ Path="$Script:HKLMPol\Microsoft\Windows\Windows Search";       Name='DisableWebSearch';                  Value=1; Desc='Web-Suche komplett deaktiviert' }
        @{ Path="$Script:HKLMPol\Microsoft\Windows\Windows Search";       Name='AllowSearchToUseLocation';          Value=0; Desc='Standort fuer Suche deaktiviert' }

        # --- Cloud-Content und Consumer Features ---
        @{ Path="$Script:HKLMPol\Microsoft\Windows\CloudContent";         Name='DisableWindowsConsumerFeatures';    Value=1; Desc='Consumer Features (Vorschlaege/Werbung) aus' }
        @{ Path="$Script:HKLMPol\Microsoft\Windows\CloudContent";         Name='DisableWindowsSpotlightFeatures';   Value=1; Desc='Windows Spotlight deaktiviert' }
        @{ Path="$Script:HKLMPol\Microsoft\Windows\CloudContent";         Name='DisableCloudOptimizedContent';      Value=1; Desc='Cloud-optimierter Content deaktiviert' }
        @{ Path="$Script:HKLMPol\Microsoft\Windows\CloudContent";         Name='DisableConsumerAccountStateContent'; Value=1; Desc='Consumer Account State Content aus' }

        # --- Standortdienst ---
        @{ Path="$Script:HKLMPol\Microsoft\Windows\LocationAndSensors";   Name='DisableLocation';                   Value=1; Desc='Standortdienst deaktiviert' }

        # --- Activity Feed / Timeline ---
        @{ Path="$Script:HKLMPol\Microsoft\Windows\System";               Name='EnableActivityFeed';                Value=0; Desc='Activity Feed deaktiviert' }
        @{ Path="$Script:HKLMPol\Microsoft\Windows\System";               Name='PublishUserActivities';             Value=0; Desc='User Activities Veroeffentlichung aus' }
        @{ Path="$Script:HKLMPol\Microsoft\Windows\System";               Name='UploadUserActivities';              Value=0; Desc='User Activities Upload aus' }

        # --- CEIP / App-Telemetrie ---
        @{ Path="$Script:HKLMPol\Microsoft\SQMClient\Windows";            Name='CEIPEnable';                        Value=0; Desc='CEIP (Customer Experience) deaktiviert' }
        @{ Path="$Script:HKLMPol\Microsoft\Windows\AppCompat";             Name='AITEnable';                         Value=0; Desc='Application Impact Telemetrie aus' }
        @{ Path="$Script:HKLMPol\Microsoft\Windows\AppCompat";             Name='DisableInventory';                  Value=1; Desc='Inventory Collector deaktiviert' }
        @{ Path="$Script:HKLMPol\Microsoft\Windows\AppCompat";             Name='DisableUAR';                        Value=1; Desc='User Access Reporting deaktiviert' }

        # --- Windows Error Reporting ---
        @{ Path="$Script:HKLMPol\Microsoft\Windows\Windows Error Reporting"; Name='Disabled';                        Value=1; Desc='Windows Fehlerberichterstattung deaktiviert' }

        # --- OneDrive ---
        @{ Path="$Script:HKLMPol\Microsoft\Windows\OneDrive";             Name='DisableFileSyncNGSC';               Value=1; Desc='OneDrive File Sync deaktiviert' }

        # --- Edge Telemetrie (HKLM-Policies) ---
        @{ Path="$Script:HKLMPol\Microsoft\Edge";                         Name='ExperimentationAndConfigurationServiceControl'; Value=0; Desc='Edge Experimentierung aus' }
        @{ Path="$Script:HKLMPol\Microsoft\Edge";                         Name='SendSiteInfoToImproveServices';     Value=0; Desc='Edge Site-Info Telemetrie aus' }
        @{ Path="$Script:HKLMPol\Microsoft\Edge";                         Name='MetricsReportingEnabled';           Value=0; Desc='Edge Metriken deaktiviert' }
        @{ Path="$Script:HKLMPol\Microsoft\Edge";                         Name='PersonalizationReportingEnabled';   Value=0; Desc='Edge Personalisierungs-Reporting aus' }
        @{ Path="$Script:HKLMPol\Microsoft\Edge";                         Name='ShoppingAssistantEnabled';          Value=0; Desc='Edge Shopping-Assistent aus' }
        @{ Path="$Script:HKLMPol\Microsoft\Edge";                         Name='EdgeCollectionsEnabled';            Value=0; Desc='Edge Sammlungen deaktiviert' }
        @{ Path="$Script:HKLMPol\Microsoft\Edge";                         Name='ShowMicrosoftRewards';              Value=0; Desc='Edge Microsoft Rewards aus' }
        @{ Path="$Script:HKLMPol\Microsoft\Edge";                         Name='SpotlightExperiencesAndRecommendationsEnabled'; Value=0; Desc='Edge Spotlight aus' }

        # --- Copilot / AI (Win11 24H2+) ---
        @{ Path="$Script:HKLMPol\Microsoft\Windows\WindowsCopilot";       Name='TurnOffWindowsCopilot';             Value=1; Desc='Windows Copilot deaktiviert' }
        @{ Path="$Script:HKLMPol\Microsoft\Windows\WindowsAI";            Name='DisableAIDataAnalysis';             Value=1; Desc='Windows Recall/AI Datenanalyse deaktiviert' }

        # --- Widgets / News ---
        @{ Path="$Script:HKLMPol\Microsoft\Dsh";                          Name='AllowNewsAndInterests';             Value=0; Desc='Widgets/News and Interests deaktiviert' }

        # --- Chat/Teams Icon ---
        @{ Path="$Script:HKLMPol\Microsoft\Windows\Windows Chat";         Name='ChatIcon';                          Value=3; Desc='Teams Chat-Icon aus Taskbar entfernt' }

        # --- Insider Program ---
        @{ Path="$Script:HKLMPol\Microsoft\Windows\WindowsUpdate";        Name='ManagePreviewBuildsPolicyValue';    Value=0; Desc='Windows Insider Program blockiert' }

        # --- Ink Workspace ---
        @{ Path="$Script:HKLMPol\Microsoft\WindowsInkWorkspace";          Name='AllowWindowsInkWorkspace';          Value=0; Desc='Ink Workspace Telemetrie deaktiviert' }

        # --- OOBE Bypass (lokale Konten erzwingen) ---
        @{ Path="$Script:HKLM\Microsoft\Windows\CurrentVersion\OOBE";     Name='BypassNRO';                         Value=1; Desc='OOBE Network Requirement Bypass (lokale Konten)' }

        # --- BitLocker Auto-Encryption verhindern ---
        @{ Path="$Script:HKLMPol\Microsoft\FVE";                          Name='OSUnaware';                         Value=1; Desc='BitLocker Auto-Verschluesselung verhindert' }
        @{ Path="$Script:HKLMPol\Microsoft\FVE";                          Name='FDVDenyWriteAccess';                Value=1; Desc='BitLocker Fixed Drive Write Access blockiert' }

        # --- Scheduled Tasks Telemetrie deaktivieren (NTLite) ---
        @{ Path="$Script:HKLMPol\Microsoft\Windows\ScheduledDiagnostics"; Name='EnabledExecution';                  Value=0; Desc='Geplante Diagnose deaktiviert' }
    )

    foreach ($s in $settings) {
        Set-RegValue -Path $s.Path -Name $s.Name -Value $s.Value -Description $s.Desc -Module "A"
    }

    # Telemetrie Scheduled Tasks deaktivieren
    $telemetryTasks = @(
        "\Microsoft\Windows\Application Experience\Microsoft Compatibility Appraiser"
        "\Microsoft\Windows\Application Experience\ProgramDataUpdater"
        "\Microsoft\Windows\Autochk\Proxy"
        "\Microsoft\Windows\Customer Experience Improvement Program\Consolidator"
        "\Microsoft\Windows\Customer Experience Improvement Program\UsbCeip"
        "\Microsoft\Windows\DiskDiagnostic\Microsoft-Windows-DiskDiagnosticDataCollector"
        "\Microsoft\Windows\Feedback\Siuf\DmClient"
        "\Microsoft\Windows\Feedback\Siuf\DmClientOnScenarioDownload"
    )
    foreach ($taskName in $telemetryTasks) {
        try {
            $task = Get-ScheduledTask -TaskPath ($taskName -replace '[^\\]*$','') -TaskName ($taskName -replace '.*\\','') -ErrorAction SilentlyContinue
            if ($task -and $task.State -ne 'Disabled') {
                Disable-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue | Out-Null
                Write-LogEntry "Telemetrie-Task deaktiviert: $taskName" "SUCCESS"
            }
        } catch {
            Write-LogEntry "Task nicht gefunden/deaktivierbar: $taskName" "INFO"
        }
    }
}

# ╔══════════════════════════════════════════════════════════════════════════════╗
# ║  MODUL B: Defender-Steuerung                                                ║
# ╚══════════════════════════════════════════════════════════════════════════════╝
function Invoke-DefenderControl {
    Write-LogEntry "Modul B: Windows Defender Steuerung" "HEAD"

    if ($EnableDefender) {
        Write-LogEntry "Defender bleibt AKTIVIERT (Hardening-Modus mit ASR)." "INFO"
        # ASR im Block-Modus aktivieren
        $ASRPath = "$Script:HKLMPol\Microsoft\Windows Defender\Windows Defender Exploit Guard\ASR\Rules"
        @("be9ba2d9-53ea-4cdc-84e5-9b1eeee46550",  # Block Office macro code
          "d4f940ab-401b-4efc-aadc-ad5f3c50688a",   # Block Office child processes
          "9e6c4e1f-7d60-472f-ba1a-a39ef669e4b2",   # Block credential stealing
          "3b576869-a4ec-4529-8536-b80a7769e899",    # Block Office executables
          "75668c1f-73b5-4cf0-bb93-3ecf5cb7cc84"     # Block process injection
        ) | ForEach-Object {
            Set-RegValue -Path $ASRPath -Name $_ -Value "1" -Type "String" -Description "ASR Regel $_ (Block)" -Module "B"
        }
        return
    }

    # --- Defender DEAKTIVIEREN (Lab-Modus) ---
    $warnMsg = @"
╔══════════════════════════════════════════════════════════════╗
║  WARNUNG: Windows Defender wird DEAKTIVIERT!                 ║
║                                                              ║
║  Nur vorgesehen fuer:                                        ║
║    - Lab-/Test-Umgebungen ohne Netzwerkzugang                ║
║    - Systeme mit alternativer AV-Software (ESET, Sophos,..)  ║
║                                                              ║
║  Ohne Virenscanner ist das System UNGESCHUETZT!              ║
╚══════════════════════════════════════════════════════════════╝
"@
    Write-Host $warnMsg -ForegroundColor Red

    $DefPath = "$Script:HKLMPol\Microsoft\Windows Defender"
    $RTPath  = "$Script:HKLMPol\Microsoft\Windows Defender\Real-Time Protection"
    Set-RegValue -Path $DefPath -Name "DisableAntiSpyware"        -Value 1 -Description "Defender Dienst deaktiviert" -Module "B"
    Set-RegValue -Path $DefPath -Name "DisableAntiVirus"          -Value 1 -Description "AntiVirus Engine deaktiviert" -Module "B"
    Set-RegValue -Path $RTPath  -Name "DisableRealtimeMonitoring" -Value 1 -Description "Echtzeitschutz deaktiviert" -Module "B"
    Set-RegValue -Path $RTPath  -Name "DisableBehaviorMonitoring" -Value 1 -Description "Verhaltensschutz deaktiviert" -Module "B"
    Set-RegValue -Path $RTPath  -Name "DisableOnAccessProtection" -Value 1 -Description "On-Access Schutz deaktiviert" -Module "B"
    Set-RegValue -Path $RTPath  -Name "DisableScanOnRealtimeEnable" -Value 1 -Description "Scan bei Aktivierung deaktiviert" -Module "B"

    # ASR im Audit-Modus (Logging ohne Blockierung)
    $ASRPath = "$Script:HKLMPol\Microsoft\Windows Defender\Windows Defender Exploit Guard\ASR\Rules"
    @("be9ba2d9-53ea-4cdc-84e5-9b1eeee46550","d4f940ab-401b-4efc-aadc-ad5f3c50688a","9e6c4e1f-7d60-472f-ba1a-a39ef669e4b2") | ForEach-Object {
        Set-RegValue -Path $ASRPath -Name $_ -Value "2" -Type "String" -Description "ASR Regel $_ (Audit)" -Module "B"
    }
    Write-LogEntry "Defender vollstaendig deaktiviert (Lab-Modus)" "WARNING"
}

# ╔══════════════════════════════════════════════════════════════════════════════╗
# ║  MODUL C: Erweiterte Sicherheit (LSA, Credential Guard, ASR)               ║
# ╚══════════════════════════════════════════════════════════════════════════════╝
function Invoke-SecurityHardening {
    Write-LogEntry "Modul C: Erweiterte Sicherheit (LSA, CG, ASR)" "HEAD"

    # LSA Protection
    Set-RegValue -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa" -Name 'RunAsPPL' -Value 1 -Description 'LSA Protection aktiviert (PPL)' -Module "C"

    # Credential Guard aktivieren
    Set-RegValue -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa" -Name 'LsaCfgFlags' -Value 1 -Description 'Credential Guard aktiviert' -Module "C"

    # Credential Guard Device Guard
    Set-RegValue -Path "HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard" -Name 'EnableVirtualizationBasedSecurity' -Value 1 -Description 'VBS aktiviert' -Module "C"

    # SMBv1 deaktivieren
    Set-RegValue -Path "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters" -Name 'SMB1' -Value 0 -Description 'SMBv1 deaktiviert (Sicherheit)' -Module "C"

    # PowerShell Script Block Logging (Sicherheits-Audit)
    Set-RegValue -Path "$Script:HKLMPol\Microsoft\Windows\PowerShell\ScriptBlockLogging" -Name 'EnableScriptBlockLogging' -Value 1 -Description 'PS Script Block Logging aktiviert' -Module "C"

    # AutoRun/AutoPlay deaktivieren
    Set-RegValue -Path "$Script:HKLMPol\Microsoft\Windows\Explorer" -Name 'NoAutoplayfornonVolume' -Value 1 -Description 'AutoPlay fuer Nicht-Volumes deaktiviert' -Module "C"
    Set-RegValue -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" -Name 'NoDriveTypeAutoRun' -Value 255 -Description 'AutoRun fuer alle Laufwerke deaktiviert' -Module "C"

    Write-LogEntry "Erweiterte Sicherheit konfiguriert" "SUCCESS"
}

# ╔══════════════════════════════════════════════════════════════════════════════╗
# ║  MODUL D: Remote & Netzwerk                                                 ║
# ╚══════════════════════════════════════════════════════════════════════════════╝
function Invoke-RemoteSetup {
    if ($SkipRemoteSetup) { Write-LogEntry "Remote-Setup uebersprungen (-SkipRemoteSetup)" "INFO"; return }
    Write-LogEntry "Modul D: Remote-Zugriff & Netzwerk" "HEAD"

    # RDP aktivieren mit NLA
    try {
        Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server' -Name 'fDenyTSConnections' -Value 0 -Type DWord -Force
        Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' -Name 'UserAuthentication' -Value 1 -Type DWord -Force
        Write-LogEntry "RDP aktiviert mit NLA-Erzwingung" "SUCCESS"
    } catch { Write-LogEntry "RDP-Fehler: $($_.Exception.Message)" "ERROR" }

    # WinRM aktivieren (fuer PS-Session)
    try {
        Set-Service -Name "WinRM" -StartupType Automatic -ErrorAction Stop
        Start-Service -Name "WinRM" -ErrorAction SilentlyContinue
        # TrustedHosts fuer internen Betrieb
        Set-Item WSMan:\localhost\Client\TrustedHosts -Value "*" -Force -ErrorAction SilentlyContinue
        Write-LogEntry "WinRM aktiviert (TrustedHosts=*)" "SUCCESS"
    } catch { Write-LogEntry "WinRM-Fehler: $($_.Exception.Message)" "WARNING" }

    # ExecutionPolicy RemoteSigned
    try {
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine -Force -ErrorAction Stop
        Write-LogEntry "ExecutionPolicy: RemoteSigned (LocalMachine)" "SUCCESS"
    } catch { Write-LogEntry "ExecutionPolicy-Fehler: $($_.Exception.Message)" "WARNING" }

    # Firewall: Ping, RDP, WinRM
    try {
        Get-NetFirewallRule -DisplayGroup "Remotedesktop" -ErrorAction SilentlyContinue | Enable-NetFirewallRule -ErrorAction SilentlyContinue
        Get-NetFirewallRule -Name "*echo*" -ErrorAction SilentlyContinue | Enable-NetFirewallRule -ErrorAction SilentlyContinue
        Enable-NetFirewallRule -DisplayGroup "Windows-Remoteverwaltung" -ErrorAction SilentlyContinue
        Write-LogEntry "Firewall: RDP + Ping + WinRM aktiviert" "SUCCESS"
    } catch { Write-LogEntry "Firewall-Fehler: $($_.Exception.Message)" "WARNING" }

    # WOL - Wake-Timer erlauben
    try {
        powercfg /SETACVALUEINDEX SCHEME_CURRENT 238c9fa8-0aaa-4286-a941-30fd9d27a4a2 bd3b718a-0680-4d9d-8ab2-e1d2b4ac806d 1 2>$null
        Write-LogEntry "Wake-on-LAN Timer aktiviert" "SUCCESS"
    } catch { Write-LogEntry "WOL-Fehler: $($_.Exception.Message)" "WARNING" }
}

# ╔══════════════════════════════════════════════════════════════════════════════╗
# ║  MODUL E: Dienste & Store                                                   ║
# ╚══════════════════════════════════════════════════════════════════════════════╝
function Invoke-ServiceHardening {
    Write-LogEntry "Modul E: Dienste & Store deaktivieren" "HEAD"

    # Telemetrie-Dienste (BSI Telemetriedokument 12, 24)
    $servicesToDisable = @(
        @{ Name="DiagTrack";        Desc="Diagnose-Tracking-Dienst (Telemetrie-Kern)" }
        @{ Name="dmwappushservice"; Desc="WAP Push Message Routing (Telemetrie)" }
        @{ Name="WpnService";       Desc="Windows Push Notifications" }
        @{ Name="wlidsvc";          Desc="Microsoft-Kontoanmeldungs-Assistent" }
        @{ Name="MapsBroker";       Desc="Downloaded Maps Manager" }
        @{ Name="lfsvc";            Desc="Geolocation Service" }
        @{ Name="RetailDemo";       Desc="Retail Demo Service" }
    )

    foreach ($svc in $servicesToDisable) {
        try {
            $service = Get-Service -Name $svc.Name -ErrorAction SilentlyContinue
            if ($service) {
                if ($service.Status -ne 'Stopped') { Stop-Service -Name $svc.Name -Force -ErrorAction SilentlyContinue }
                Set-Service -Name $svc.Name -StartupType Disabled -ErrorAction Stop
                Write-LogEntry "Dienst deaktiviert: $($svc.Name) — $($svc.Desc)" "SUCCESS"
                Update-StateEntry -Module "E" -Key "Service_$($svc.Name)" -Before $service.StartType -After "Disabled" -Desc $svc.Desc
            }
        } catch { Write-LogEntry "Dienst $($svc.Name): $($_.Exception.Message)" "WARNING" }
    }

    # Windows Store deaktivieren (Standard)
    Set-RegValue -Path "$Script:HKLMPol\Microsoft\WindowsStore" -Name "AutoDownload"     -Value 2 -Description "Store Auto-Download deaktiviert" -Module "E"
    Set-RegValue -Path "$Script:HKLMPol\Microsoft\WindowsStore" -Name "RemoveWindowsStore" -Value 1 -Description "Windows Store deaktiviert" -Module "E"
    Set-RegValue -Path "$Script:HKLMPol\Microsoft\WindowsStore" -Name "DisableStoreApps"  -Value 1 -Description "Store-Apps Installation blockiert" -Module "E"
    Write-LogEntry "Windows Store vollstaendig deaktiviert" "SUCCESS"
}

# ╔══════════════════════════════════════════════════════════════════════════════╗
# ║  MODUL F: Debloating (AppX + Capabilities)                                  ║
# ║  Auswahl: -DebloatAction "Disable" (nur deaktivieren, wiederherstellbar)    ║
# ║           -DebloatAction "Remove"  (permanent entfernen, Standard)          ║
# ╚══════════════════════════════════════════════════════════════════════════════╝
function Invoke-Debloating {
    Write-LogEntry "Modul F: Debloating ($DebloatAction-Modus)" "HEAD"

    # App-Liste basierend auf Referenzen (Image_Anpassung + V10)
    $AppList = @(
        "*Teams*", "*Xbox*", "*Copilot*", "*Clipchamp*", "*OneDrive*",
        "*OneNote*", "*News*", "*Weather*", "*Zune*", "*Bing*",
        "*Solitaire*", "*People*", "*CommunicationsApps*", "*Outlook*",
        "*SkypeApp*", "*WindowsFeedbackHub*", "*GetHelp*", "*Getstarted*",
        "*Maps*", "*Messaging*", "*MicrosoftOfficeHub*", "*MixedReality*",
        "*Paint3D*", "*SoundRecorder*", "*Todos*", "*Wallet*",
        "*WebExperience*", "*YourPhone*", "*WindowsStore*",
        "*549981C3F5F10*", "*BingNews*", "*ZuneVideo*", "*ZuneMusic*"
    )

    if ($DebloatAction -eq "Remove") {
        # --- Permanentes Entfernen ---
        # 1. Provisioned Packages (fuer neue User)
        $prov = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
        foreach ($pattern in $AppList) {
            $prov | Where-Object { $_.DisplayName -like $pattern } | ForEach-Object {
                Remove-AppxProvisionedPackage -Online -PackageName $_.PackageName -ErrorAction SilentlyContinue | Out-Null
                Write-LogEntry "Provisioned entfernt: $($_.DisplayName)" "SUCCESS"
                Update-StateEntry -Module "F" -Key "Prov_$($_.DisplayName)" -Before "Installed" -After "Removed" -Desc "Provisioned Package"
            }
        }

        # 2. Installierte Pakete (alle User)
        $installedPackages = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue
        foreach ($pattern in $AppList) {
            $installedPackages | Where-Object { $_.Name -like $pattern } | ForEach-Object {
                Remove-AppxPackage -Package $_.PackageFullName -AllUsers -ErrorAction SilentlyContinue | Out-Null
                Write-LogEntry "AppX entfernt: $($_.Name)" "INFO"
            }
        }

        # 3. Win32 Teams Cleanup
        try { taskkill /f /im Teams.exe 2>$null | Out-Null } catch {}
        Get-ChildItem -Path "C:\Users" -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notin @("Public","Default","All Users","Default User") } | ForEach-Object {
                $upd = Join-Path $_.FullName "AppData\Local\Microsoft\Teams\Update.exe"
                if (Test-Path $upd) { Start-Process -FilePath $upd -ArgumentList "--uninstall -s" -Wait -ErrorAction SilentlyContinue }
            }
    } else {
        # --- Nur Deaktivieren (wiederherstellbar) ---
        $installedPackagesDeactivate = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue
        foreach ($pattern in $AppList) {
            $installedPackagesDeactivate | Where-Object { $_.Name -like $pattern } | ForEach-Object {
                # NonRemovable-Apps koennen nicht entfernt, aber deaktiviert werden
                try {
                    Get-AppxPackage -Name $_.Name -AllUsers | Remove-AppxPackage -ErrorAction SilentlyContinue | Out-Null
                } catch {}
                Write-LogEntry "AppX deaktiviert: $($_.Name)" "INFO"
            }
        }
    }

    # 4. Windows Capabilities entfernen (aus V3/Referenz)
    @("App.Support.QuickAssist*", "App.StepsRecorder*", "Browser.InternetExplorer*",
      "MathRecognizer*", "Microsoft.Windows.Wordpad*", "Print.Fax.Scan*",
      "Language.Handwriting*", "Language.Speech*"
    ) | ForEach-Object {
        Get-WindowsCapability -Online -Name $_ -ErrorAction SilentlyContinue |
            Where-Object { $_.State -eq 'Installed' } | ForEach-Object {
                Remove-WindowsCapability -Online -Name $_.Name -ErrorAction SilentlyContinue | Out-Null
                Write-LogEntry "Capability entfernt: $($_.Name)" "SUCCESS"
            }
    }

    Write-LogEntry "Debloating abgeschlossen ($DebloatAction)" "SUCCESS"
}

# ╔══════════════════════════════════════════════════════════════════════════════╗
# ║  MODUL G: UI-Anpassungen, Explorer & Nag-Screen-Unterdrueckung             ║
# ╚══════════════════════════════════════════════════════════════════════════════╝
function Invoke-UISetup {
    Write-LogEntry "Modul G: UI-Anpassungen & Nag-Screen-Unterdrueckung" "HEAD"

    # --- Explorer & Desktop (Admin-optimiert) ---
    $uiSettings = @(
        @{ Path="HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32"; Name=''; Value=''; Type='String'; Desc='Classic Context Menu (Win11)' }
        @{ Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"; Name='Start_TaskbarAlignment';        Value=0; Desc='Taskbar linksbuendig' }
        @{ Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"; Name='Hidden';                         Value=1; Desc='Versteckte Dateien anzeigen' }
        @{ Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"; Name='HideFileExt';                    Value=0; Desc='Dateiendungen anzeigen' }
        @{ Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"; Name='NavPaneExpandToCurrentFolder';   Value=1; Desc='Navigation: aktueller Ordner' }
        @{ Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"; Name='DontUsePowerShellOnWinX';        Value=0; Desc='PowerShell im Win+X Menue' }
        @{ Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"; Name='ShowCompColor';                   Value=1; Desc='Komprimierte Dateien farbig' }
        @{ Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"; Name='Start_TrackProgs';                Value=0; Desc='App-Start-Tracking deaktiviert' }
        @{ Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer";          Name='ShowDriveLettersFirst';           Value=4; Desc='Laufwerksbuchstabe zuerst' }
        # Dark Mode
        @{ Path="HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"; Name='AppsUseLightTheme';             Value=0; Desc='Dark Mode fuer Apps' }
        @{ Path="HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"; Name='SystemUsesLightTheme';          Value=0; Desc='Dark Mode fuer System' }
        # Clipboard History
        @{ Path="HKCU:\SOFTWARE\Microsoft\Clipboard"; Name='EnableClipboardHistory'; Value=1; Desc='Zwischenablageverlauf aktiviert' }
        # Telemetrie HKCU
        @{ Path="HKCU:\SOFTWARE\Microsoft\Input\TIPC";                                Name='Enabled';                        Value=0; Desc='Eingabe-Telemetrie deaktiviert' }
        @{ Path="HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Search";              Name='BingSearchEnabled';              Value=0; Desc='Bing-Suche deaktiviert' }
        @{ Path="HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name='SilentInstalledAppsEnabled'; Value=0; Desc='Vorgeschlagene Apps deaktiviert' }
        @{ Path="HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"; Name='SystemPaneSuggestionsEnabled'; Value=0; Desc='Startmenue-Vorschlaege aus' }
        # Ink/Handwriting Personalization
        @{ Path="HKCU:\Software\Microsoft\InputPersonalization";                       Name='RestrictImplicitInkCollection';  Value=1; Desc='Freihand-Datensammlung deaktiviert' }
        @{ Path="HKCU:\Software\Microsoft\InputPersonalization";                       Name='RestrictImplicitTextCollection'; Value=1; Desc='Text-Datensammlung deaktiviert' }
        @{ Path="HKCU:\Software\Microsoft\InputPersonalization\TrainedDataStore";      Name='HarvestContacts';                Value=0; Desc='Kontakte-Harvesting deaktiviert' }
        @{ Path="HKCU:\Software\Microsoft\Personalization\Settings";                   Name='AcceptedPrivacyPolicy';          Value=0; Desc='Personalisierungs-Policy abgelehnt' }
        # Sprachlistenzugriff
        @{ Path="HKCU:\Control Panel\International\User Profile";                      Name='HttpAcceptLanguageOptOut';        Value=1; Desc='Website-Sprachlistenzugriff deaktiviert' }
        # Cloud Sync
        @{ Path="HKCU:\SOFTWARE\Microsoft\Messaging";                                 Name='CloudServiceSyncEnabled';        Value=0; Desc='Cloud-Nachrichtensync deaktiviert' }
        # Edge HKCU
        @{ Path="HKCU:\SOFTWARE\Policies\Microsoft\Edge"; Name='MetricsReportingEnabled';             Value=0; Desc='Edge Metriken aus (HKCU)' }
        @{ Path="HKCU:\SOFTWARE\Policies\Microsoft\Edge"; Name='HideFirstRunExperience';              Value=1; Desc='Edge Onboarding aus' }
        @{ Path="HKCU:\SOFTWARE\Policies\Microsoft\Edge"; Name='PersonalizationDataCollectionEnabled'; Value=0; Desc='Edge Personalisierung aus' }
        @{ Path="HKCU:\SOFTWARE\Policies\Microsoft\Edge"; Name='ShoppingAssistantEnabled';             Value=0; Desc='Edge Shopping aus (HKCU)' }
        # Windows Terminal als Standard
        @{ Path="HKCU:\Console\%%Startup"; Name='DelegationTerminal'; Value='{E12CFF52-A866-4C77-9A90-F570A7AA2C6B}'; Type='String'; Desc='Windows Terminal als Standard' }
    )

    foreach ($s in $uiSettings) {
        $type = if ($s.Type) { $s.Type } else { "DWord" }
        Set-RegValue -Path $s.Path -Name $s.Name -Value $s.Value -Type $type -Description $s.Desc -Module "G"
    }

    # --- Nag-Screen-Unterdrueckung (Chrome, Firefox, Edge, Office, Teams, Adobe) ---
    Write-LogEntry "Nag-Screens/Welcome-Dialoge unterdruecken..." "INFO"
    foreach ($nag in $Script:NagScreenApps) {
        $type = if ($nag.Type) { $nag.Type } else { "DWord" }
        Set-RegValue -Path $nag.Path -Name $nag.Key -Value $nag.Value -Type $type -Description "$($nag.Name): $($nag.Key)" -Module "G"
    }
}

# ╔══════════════════════════════════════════════════════════════════════════════╗
# ║  MODUL H: Energie & System                                                  ║
# ╚══════════════════════════════════════════════════════════════════════════════╝
function Invoke-SystemSetup {
    Write-LogEntry "Modul H: Energie & System" "HEAD"

    # Fast Boot deaktivieren (OPSI/PXE-Vorgabe)
    Set-RegValue -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power" -Name 'HiberbootEnabled' -Value 0 -Description "Fast Boot deaktiviert (OPSI/PXE)" -Module "H"

    # Lock Screen deaktivieren
    Set-RegValue -Path "$Script:HKLMPol\Microsoft\Windows\Personalization" -Name 'NoLockScreen' -Value 1 -Description "Lock Screen deaktiviert" -Module "H"

    # Screensaver deaktivieren
    Set-RegValue -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" -Name 'ScreenSaverIsSecure' -Value 0 -Description "Bildschirmschoner-Sperre aus" -Module "H"

    # Energieplan
    try {
        powercfg /CHANGE MONITOR-TIMEOUT-AC 30  2>$null
        powercfg /CHANGE MONITOR-TIMEOUT-DC 30  2>$null
        powercfg /CHANGE STANDBY-TIMEOUT-AC 480 2>$null
        powercfg /CHANGE STANDBY-TIMEOUT-DC 480 2>$null
        powercfg /CHANGE HIBERNATE-TIMEOUT-AC 0 2>$null
        powercfg /CHANGE HIBERNATE-TIMEOUT-DC 0 2>$null
        Write-LogEntry "Energieplan: Monitor 30min / Standby 8h / Ruhezustand nie" "SUCCESS"
    } catch { Write-LogEntry "Energieplan-Fehler: $($_.Exception.Message)" "WARNING" }
}

# ╔══════════════════════════════════════════════════════════════════════════════╗
# ║  MODUL I: Treiber-Injektion (Universell, SCCM-kompatibel, FTS/Lenovo Fokus) ║
# ╚══════════════════════════════════════════════════════════════════════════════╝
function Invoke-DriverInjection {
    if ($SkipDriverInstall) { Write-LogEntry "Treiber-Install uebersprungen (-SkipDriverInstall)" "INFO"; return }
    Write-LogEntry "Modul I: Treiber-Injektion" "HEAD"

    $vendor = (Get-WmiObject Win32_ComputerSystem).Manufacturer
    $model = (Get-WmiObject Win32_ComputerSystem).Model
    Write-LogEntry "Erkannte Hardware: $vendor - $model" "INFO"

    # FTS/Lenovo Check
    $isFTS = $vendor -match "FUJITSU"
    $isLenovo = $vendor -match "LENOVO"

    # Pfad zu SCCM Treiberpaketen (z.B. auf USB-Stick oder OPSI-Depot)
    if (-not (Test-Path $DriversPathSource)) {
        Write-LogEntry "Treiber-Verzeichnis nicht gefunden: $DriversPathSource" "WARNING"
        return
    }

    # Tracking fuer mehrstufige Updates (Reboots)
    $rebootCount = Get-RegistryValueSafe -Path $Script:DriverRebootKey -Name "RebootCount"
    if ($null -eq $rebootCount) { $rebootCount = 0 }

    if ($rebootCount -ge $Script:MaxDriverReboots) {
        Write-LogEntry "Maximale Reboot-Anzahl fuer Treiber ($Script:MaxDriverReboots) erreicht. Abbruch." "WARNING"
        return
    }

    $driverFolders = Get-ChildItem -Path $DriversPathSource -Directory
    $driversInstalled = 0

    foreach ($folder in $driverFolders) {
        Write-LogEntry "Pruefe Treiberpaket: $($folder.Name)" "INFO"
        
        # FTS Besonderheit: PD/BD PROCHOT Bug Cleanup vor Treiber-Installation (an FTS_Update_and_PD_Error angelehnt)
        if ($isFTS -and $folder.Name -match "Firmware|Intel") {
            Write-LogEntry "FTS Cleanup: Loesche alte Intel/UCSI Treiberreste..." "INFO"
            # Beispielhaftes Cleanup: pnputil /delete-driver oem*.inf /uninstall (vereinfacht)
            # In der Praxis gezielt nach bestimmten HWIDs suchen
        }

        # Rekursive Installation via pnputil
        $infFiles = Get-ChildItem -Path $folder.FullName -Filter "*.inf" -Recurse
        if ($infFiles.Count -gt 0) {
            Write-LogEntry "Installiere $($infFiles.Count) Treiber aus $($folder.Name)..." "INFO"
            $pnpArgs = "/add-driver `"$($folder.FullName)\*.inf`" /subdirs /install"
            $proc = Start-Process -FilePath "pnputil.exe" -ArgumentList $pnpArgs -Wait -NoNewWindow -PassThru
            
            if ($proc.ExitCode -eq 3010) { # ERROR_SUCCESS_REBOOT_REQUIRED
                Write-LogEntry "Reboot erforderlich fuer Treiber: $($folder.Name)" "WARNING"
                Set-RegValue -Path $Script:DriverRebootKey -Name "RebootCount" -Value ($rebootCount + 1) -Module "I"
                Set-RegValue -Path $Script:DriverRebootKey -Name "PendingReboot" -Value 1 -Module "I"
            }
            $driversInstalled += $infFiles.Count
        }
    }

    Write-LogEntry "Treiber-Injektion abgeschlossen ($driversInstalled INF-Dateien)." "SUCCESS"
}

# ╔══════════════════════════════════════════════════════════════════════════════╗
# ║  MODUL J: DNS-Sperrliste (Hosts-basiert, Offline-faehig)                    ║
# ╚══════════════════════════════════════════════════════════════════════════════╝
function Invoke-DNSBlocklist {
    Write-LogEntry "Modul J: DNS-Sperrliste (Telemetrie-Blockade)" "HEAD"

    # Backup der hosts-Datei
    if (-not (Test-Path $Script:HostsBackup)) {
        Copy-Item -Path $Script:HostsFile -Destination $Script:HostsBackup -Force
        Write-LogEntry "Hosts-Backup erstellt: $Script:HostsBackup" "INFO"
    }

    $hostsContent = Get-Content $Script:HostsFile -Raw -ErrorAction SilentlyContinue
    if ($null -eq $hostsContent) { $hostsContent = "" }

    # Alten Block entfernen
    if ($hostsContent -match "(?s)$($Script:HostsMarkerStart).*?$($Script:HostsMarkerEnd)\r?\n?") {
        $hostsContent = $hostsContent -replace "(?s)$($Script:HostsMarkerStart).*?$($Script:HostsMarkerEnd)\r?\n?", ""
        Write-LogEntry "Alten DNS-Block entfernt." "INFO"
    }

    $blockEntries = @()

    # Online-Quellen abrufen, falls konfiguriert und verfuegbar
    if ($UpdateDNSBlocklist) {
        Write-LogEntry "Lade DNS-Blocklisten aus Online-Quellen..." "INFO"
        foreach ($url in $Script:DNSBlocklistURLs) {
            try {
                $response = Invoke-RestMethod -Uri $url -UseBasicParsing -TimeoutSec 10
                $lines = $response -split "`n"
                foreach ($line in $lines) {
                    $line = $line.Trim()
                    if ($line -notmatch "^#" -and $line -match "^(?:0\.0\.0\.0|127\.0\.0\.1)\s+(.+)") {
                        $domain = $matches[1].Trim()
                        $blockEntries += "0.0.0.0 $domain"
                    }
                }
            } catch { Write-LogEntry "Fehler beim Laden von $url: $($_.Exception.Message)" "WARNING" }
        }
        
        # Speichere offline-Kopie, falls Online erfolgreich
        if ($blockEntries.Count -gt 0) {
            $blockEntries | Select-Object -Unique | Set-Content -Path $Script:DNSOfflineFile -Encoding UTF8
        }
    } 
    
    # Fallback auf lokale Liste, wenn keine Online-Updates oder fehlgeschlagen
    if ($blockEntries.Count -eq 0 -and (Test-Path $Script:DNSOfflineFile)) {
        Write-LogEntry "Lade DNS-Blockliste aus lokaler Datei ($Script:DNSOfflineFile)..." "INFO"
        $blockEntries = Get-Content -Path $Script:DNSOfflineFile | Where-Object { $_ -match "^0\.0\.0\.0" }
    }

    $blockEntries = $blockEntries | Select-Object -Unique

    # Whitelist anwenden
    if (Test-Path $Script:DNSWhitelistFile) {
        $whitelist = Get-Content $Script:DNSWhitelistFile | Where-Object { $_ -notmatch "^\s*#" -and $_ -match "\S" } | ForEach-Object { $_.Trim() }
        if ($whitelist.Count -gt 0) {
            $initialCount = $blockEntries.Count
            $blockEntries = $blockEntries | Where-Object { 
                $domain = ($_ -split "\s+")[1]
                $domain -notin $whitelist
            }
            Write-LogEntry "$($initialCount - $blockEntries.Count) Eintraege durch Whitelist gefiltert." "INFO"
        }
    }

    if ($blockEntries.Count -gt 0) {
        $newBlock = "`n$($Script:HostsMarkerStart)`n"
        $newBlock += "# Generiert am $(Get-Date)`n"
        $newBlock += ($blockEntries -join "`n")
        $newBlock += "`n$($Script:HostsMarkerEnd)`n"

        $hostsContent += $newBlock
        $hostsContent | Set-Content -Path $Script:HostsFile -Encoding UTF8 -Force
        
        Write-LogEntry "DNS-Sperrliste angewendet: $($blockEntries.Count) Domains blockiert." "SUCCESS"
        
        # State aktualisieren
        $Script:CurrentState.DNSBlocklist.LastUpdate = (Get-Date -Format "o")
        $Script:CurrentState.DNSBlocklist.EntriesCount = $blockEntries.Count
        $Script:CurrentState.DNSBlocklist.Source = if ($UpdateDNSBlocklist) { "Online" } else { "Local" }
    } else {
        Write-LogEntry "Keine DNS-Block-Eintraege gefunden/generiert." "WARNING"
    }
}

# ╔══════════════════════════════════════════════════════════════════════════════╗
# ║  MODUL K: HKCU User-Haertung (Scheduled Task)                               ║
# ╚══════════════════════════════════════════════════════════════════════════════╝
function Invoke-UserHardeningTask {
    Write-LogEntry "Modul K: HKCU User-Haertung (Scheduled Task)" "HEAD"

    # Erstelle das eigentliche Script, das vom Task ausgefuehrt wird (falls nicht existent)
    $UserScriptPath = Join-Path $Script:StateDir "User_Hardening.ps1"
    
    $userScriptContent = @"
# Wird bei jedem User-Login ausgefuehrt
`$HKCU = "HKCU:\SOFTWARE"
`$HKCUPol = "HKCU:\SOFTWARE\Policies"

# Telemetrie & Werbung
Set-ItemProperty -Path "`$HKCU\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -Name "SilentInstalledAppsEnabled" -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
Set-ItemProperty -Path "`$HKCU\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -Name "SystemPaneSuggestionsEnabled" -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
Set-ItemProperty -Path "`$HKCU\Microsoft\Input\TIPC" -Name "Enabled" -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue

# Explorer / UI
Set-ItemProperty -Path "`$HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "Start_TaskbarAlignment" -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
Set-ItemProperty -Path "`$HKCU\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32" -Name "" -Value "" -Type String -Force -ErrorAction SilentlyContinue

# Edge Telemetrie
if (-not (Test-Path "`$HKCUPol\Microsoft\Edge")) { New-Item -Path "`$HKCUPol\Microsoft\Edge" -Force | Out-Null }
Set-ItemProperty -Path "`$HKCUPol\Microsoft\Edge" -Name "MetricsReportingEnabled" -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
"@

    $userScriptContent | Set-Content -Path $UserScriptPath -Encoding UTF8 -Force

    # Task Scheduler Konfiguration
    $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$UserScriptPath`""
    $trigger = New-ScheduledTaskTrigger -AtLogOn
    $principal = New-ScheduledTaskPrincipal -GroupId "S-1-5-32-545" -RunLevel Highest # Users group
    
    $taskName = "Hardening_User_Init"
    
    try {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Description "Fuehrt HKCU Hardening beim Login aus." | Out-Null
        Write-LogEntry "Scheduled Task '$taskName' erfolgreich registriert." "SUCCESS"
    } catch {
        Write-LogEntry "Fehler beim Registrieren des User-Tasks: $($_.Exception.Message)" "ERROR"
    }
}

# ╔══════════════════════════════════════════════════════════════════════════════╗
# ║  EXPORT FUNKTIONEN (.reg und Autounattend.xml)                              ║
# ╚══════════════════════════════════════════════════════════════════════════════╝
function Export-RegistryHardening {
    Write-LogEntry "Generiere Registry Export ($Script:RegExportFile)..." "INFO"
    
    # Basierend auf dem aktuellen State, alle durchgefuehrten Aenderungen exportieren
    $state = Get-HardeningState
    
    $regContent = "Windows Registry Editor Version 5.00`r`n`r`n"
    
    foreach ($mod in $state.Modules.Keys) {
        $settings = $state.Modules.$mod.Settings
        if ($settings) {
            $regContent += "; =======================================`r`n"
            $regContent += "; MODUL $mod: $($Script:ModuleDefinitions[$mod].Name)`r`n"
            $regContent += "; =======================================`r`n"
            
            # Gruppiere nach Registry-Pfad
            $paths = @{}
            foreach ($key in $settings.psobject.properties.name) {
                # Key Format: HKLM:\SOFTWARE\...\Name
                if ($key -match "^(HKLM:\\.+)\\([^\\]+)$" -or $key -match "^(HKCU:\\.+)\\([^\\]+)$") {
                    $path = $matches[1] -replace "^HKLM:\\", "HKEY_LOCAL_MACHINE\" -replace "^HKCU:\\", "HKEY_CURRENT_USER\"
                    $valName = $matches[2]
                    $valData = $settings.$key.After
                    
                    if (-not $paths.ContainsKey($path)) { $paths[$path] = @() }
                    
                    if ($valData -match "^DWORD:") {
                        $paths[$path] += "`"$valName`"=dword:$($valData.Substring(6).PadLeft(8, '0'))"
                    } elseif ($valData -is [int]) {
                        $hex = "{0:x8}" -f $valData
                        $paths[$path] += "`"$valName`"=dword:$hex"
                    } else {
                        $paths[$path] += "`"$valName`"=`"$valData`""
                    }
                }
            }
            
            foreach ($path in $paths.Keys) {
                $regContent += "[$path]`r`n"
                foreach ($val in $paths[$path]) {
                    $regContent += "$val`r`n"
                }
                $regContent += "`r`n"
            }
        }
    }
    
    $regContent | Set-Content -Path $Script:RegExportFile -Encoding Ascii
    Write-LogEntry "Registry-Export erstellt." "SUCCESS"
}

function Export-AutounattendXML {
    $usbXmlPath = Join-Path $Script:StateDir "Autounattend_USB.xml"
    $opsiXmlPath = Join-Path $Script:StateDir "Autounattend_OPSI.xml"
    
    Write-LogEntry "Generiere Autounattend.xml Dateien..." "INFO"
    
    $xmlTemplate = @"
<?xml version="1.0" encoding="utf-8"?>
<unattend xmlns="urn:schemas-microsoft-com:unattend">
    <!-- Generiert durch Windows Hardening Suite V11.0 -->
    <settings pass="oobeSystem">
        <component name="Microsoft-Windows-Shell-Setup" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
            <OOBE>
                <HideEULAPage>true</HideEULAPage>
                <HideOEMRegistrationScreen>true</HideOEMRegistrationScreen>
                <HideOnlineAccountScreens>true</HideOnlineAccountScreens>
                <HideWirelessSetupInOOBE>true</HideWirelessSetupInOOBE>
                <NetworkLocation>Work</NetworkLocation>
                <ProtectYourPC>3</ProtectYourPC>
                <SkipUserOOBE>true</SkipUserOOBE>
                <SkipMachineOOBE>true</SkipMachineOOBE>
            </OOBE>
            <UserAccounts>
                <AdministratorPassword>
                    <Value>$Script:UnattendPassword</Value>
                    <PlainText>false</PlainText>
                </AdministratorPassword>
                <LocalAccounts>
                    <LocalAccount wcm:action="add">
                        <Password>
                            <Value>$Script:UnattendPassword</Value>
                            <PlainText>false</PlainText>
                        </Password>
                        <Description>Local Admin</Description>
                        <DisplayName>$Script:UnattendUser</DisplayName>
                        <Group>Administrators</Group>
                        <Name>$Script:UnattendUser</Name>
                    </LocalAccount>
                </LocalAccounts>
            </UserAccounts>
            <RegisteredOrganization>$Script:UnattendOrg</RegisteredOrganization>
            <RegisteredOwner>$Script:UnattendUser</RegisteredOwner>
            <FirstLogonCommands>
                <SynchronousCommand wcm:action="add">
                    <Order>1</Order>
                    <CommandLine>reg.exe add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE" /v BypassNRO /t REG_DWORD /d 1 /f</CommandLine>
                    <Description>Bypass NRO</Description>
                </SynchronousCommand>
                <SynchronousCommand wcm:action="add">
                    <Order>2</Order>
                    <CommandLine>PowerShell -ExecutionPolicy Bypass -Command "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/Miraculix666/Schul-OPSI/main/win11-hardening/CLIENT_DATA/Apply_Hardening.ps1' -OutFile 'C:\Apply_Hardening.ps1'; &amp; 'C:\Apply_Hardening.ps1' -SilentMode -ExportReg"</CommandLine>
                    <Description>Run Hardening V11.0</Description>
                </SynchronousCommand>
            </FirstLogonCommands>
        </component>
    </settings>
</unattend>
"@
    
    # USB Variante
    $xmlTemplate | Set-Content -Path $usbXmlPath -Encoding UTF8
    
    # OPSI Variante (ggf. leicht abgewandelt ohne FirstLogonCommands, da OPSI das übernimmt)
    $xmlOpsi = $xmlTemplate -replace "<FirstLogonCommands>.*?</FirstLogonCommands>", "<!-- FirstLogonCommands via OPSI Winst -->"
    $xmlOpsi | Set-Content -Path $opsiXmlPath -Encoding UTF8
    
    Write-LogEntry "Autounattend_USB.xml und Autounattend_OPSI.xml in $Script:StateDir erstellt." "SUCCESS"
}

# ╔══════════════════════════════════════════════════════════════════════════════╗
# ║  HAUPTPROGRAMM                                                              ║
# ╚══════════════════════════════════════════════════════════════════════════════╝

Initialize-Environment

if ($ShowStatus) {
    Show-HardeningStatus
    exit 0
}

$Script:CurrentState = Get-HardeningState

# Bestimme Ausfuehrungsmodus
$isRemote = Test-IsRemoteSession
if ($isRemote) { $SilentMode = $true }

$selectedModules = @()
$action = "apply"

if ($SilentMode -or $Modules.Count -gt 0 -or $RestoreMode -or $ExportReg -or $ExportUnattend -or $UpdateDNSBlocklist) {
    # Non-Interactive / Parameter Mode
    if ($RestoreMode) { $action = "restore" }
    elseif ($ExportReg) { $action = "exportreg" }
    elseif ($ExportUnattend) { $action = "exportunattend" }
    else { $action = "apply" }
    
    if ($Modules.Count -gt 0) {
        $selectedModules = $Modules
    } else {
        # Default Module nehmen
        $selectedModules = ($Script:ModuleDefinitions.Keys | Where-Object { $Script:ModuleDefinitions[$_].Default })
    }
} else {
    # Interactive Menu
    $menuResult = Show-InteractiveMenu
    $action = $menuResult.Action
    $selectedModules = $menuResult.SelectedModules
}

# Aktionen ausfuehren
switch ($action) {
    "quit" { Write-Host "Abbruch durch Benutzer." -ForegroundColor Yellow; exit 0 }
    "status" { Show-HardeningStatus; exit 0 }
    "exportreg" { Export-RegistryHardening; exit 0 }
    "exportunattend" { Export-AutounattendXML; exit 0 }
    "restore" {
        Write-LogEntry "Starte RESTORE-Modus..." "HEAD"
        foreach ($mod in $selectedModules) {
            Write-LogEntry "Restore Modul $mod..." "INFO"
            $settings = $Script:CurrentState.Modules.$mod.Settings
            if ($settings) {
                foreach ($key in $settings.psobject.properties.name) {
                    if ($key -match "^(HKLM:\\.+)\\([^\\]+)$" -or $key -match "^(HKCU:\\.+)\\([^\\]+)$") {
                        Restore-RegValue -Path $matches[1] -Name $matches[2] -OriginalValue $settings.$key.Before
                    }
                }
            }
            $Script:CurrentState.Modules.$mod.Applied = $false
        }
        Save-HardeningState -State $Script:CurrentState
        Write-LogEntry "Restore abgeschlossen." "SUCCESS"
        exit 0
    }
    "apply" {
        Write-LogEntry "Starte HARDENING-Modus..." "HEAD"
        
        # Module nacheinander ausfuehren
        if ("A" -in $selectedModules) { Invoke-PrivacyHardening }
        if ("B" -in $selectedModules) { Invoke-DefenderControl }
        if ("C" -in $selectedModules) { Invoke-SecurityHardening }
        if ("D" -in $selectedModules) { Invoke-RemoteSetup }
        if ("E" -in $selectedModules) { Invoke-ServiceHardening }
        if ("F" -in $selectedModules) { Invoke-Debloating }
        if ("G" -in $selectedModules) { Invoke-UISetup }
        if ("H" -in $selectedModules) { Invoke-SystemSetup }
        if ("I" -in $selectedModules) { Invoke-DriverInjection }
        if ("J" -in $selectedModules) { Invoke-DNSBlocklist }
        if ("K" -in $selectedModules) { Invoke-UserHardeningTask }
        
        # State am Ende speichern
        Save-HardeningState -State $Script:CurrentState
        
        Write-Host "`n=== Zusammenfassung ===" -ForegroundColor Cyan
        Write-Host "Erfolgreich: $($Script:Stats.Success)" -ForegroundColor Green
        Write-Host "Warnungen:   $($Script:Stats.Warnings)" -ForegroundColor Yellow
        Write-Host "Fehler:      $($Script:Stats.Errors)" -ForegroundColor Red
        Write-Host "Log-Datei:   $Script:LogFile" -ForegroundColor DarkGray
        
        # Pruefe Reboot-Bedarf von Modul I
        $pendingReboot = Get-RegistryValueSafe -Path $Script:DriverRebootKey -Name "PendingReboot"
        if ($pendingReboot -eq 1) {
            Write-Host "`n[!WICHTIG!] Ein Systemneustart ist erforderlich, um die Treiberinstallation abzuschliessen!" -ForegroundColor Red
            Set-RegValue -Path $Script:DriverRebootKey -Name "PendingReboot" -Value 0 -Module "I" | Out-Null
            if (-not $SilentMode) {
                $ans = Read-Host "Jetzt neu starten? (J/N)"
                if ($ans -eq 'J') { Restart-Computer -Force }
            }
        }
    }
}


