<#
.SYNOPSIS
    PowerSetup_Set-CustomPowerSettings.ps1 - V5.4.0 (Full RDP Transparency Edition)
.DESCRIPTION
    Uebernimmt die absolute Kontrolle ueber alle Energieprofile. 
    Beinhaltet Path-Fixes, detailliertes RDP-Session-Override und vollständige Klartext-Verifizierung.
#>

[CmdletBinding()]
param (
    [switch]$Interactive,
    [switch]$RunAsSystemMode,

    # --- SCHALTER & DECKEL ---
    [int]$LidActionAC = 1,
    [int]$LidActionDC = 1,
    [int]$PowerButtonAC = 2,
    [int]$PowerButtonDC = 2,
    [int]$SleepButtonAction = 1,

    # --- TIMEOUTS ---
    [int]$DisplayOffAC = 8,
    [int]$DisplayOffDC = 5,
    [int]$SleepAC = 45,
    [int]$SleepDC = 20,
    [int]$HibernateAC = 120,
    [int]$HibernateDC = 60,
    [int]$DiskSpindownAC = 5,
    [int]$DiskSpindownDC = 1,

    # --- SPERRBILDSCHIRM & GRACE PERIOD ---
    [int]$LockScreenTimeoutAC = 8,
    [int]$LockScreenTimeoutDC = 5,
    [int]$LockScreenGracePeriod = 30,

    # --- LEISTUNG ---
    [int]$CoolingPolicyAC = 1,
    [int]$CoolingPolicyDC = 0,
    [int]$CpuMinAC = 5,
    [int]$CpuMinDC = 5,
    [int]$CpuMaxAC = 100,
    [int]$CpuMaxDC = 100,
    [int]$WlanWirelessModeAC = 0,
    [int]$WlanWirelessModeDC = 2,
    [int]$PcieAspmAC = 0,
    [int]$PcieAspmDC = 2,

    # --- SYSTEM-OPTIONEN ---
    [switch]$EnableFastBoot,
    [switch]$ShowHibernateInMenu = $true
)

# --- INITIALISIERUNG & ELEVATION ---
[System.Threading.Thread]::CurrentThread.CurrentCulture = New-Object "System.Globalization.CultureInfo"("de-DE")
[System.Threading.Thread]::CurrentThread.CurrentUICulture = New-Object "System.Globalization.CultureInfo"("de-DE")
$VerbosePreference = "Continue"

# Bulletproof Path Detection
$GlobalScriptPath = $PSCommandPath
if ([string]::IsNullOrWhiteSpace($GlobalScriptPath)) { $GlobalScriptPath = $MyInvocation.MyCommand.Definition }

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    if ([string]::IsNullOrWhiteSpace($GlobalScriptPath) -or -not (Test-Path $GlobalScriptPath)) {
        Write-Warning "Fehler: Pfad unbekannt. Bitte das Skript erst als .ps1 speichern und per Rechtsklick oder Konsole starten!"
        Exit
    }
    $myArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$GlobalScriptPath`""
    Start-Process powershell.exe -Verb RunAs -ArgumentList $myArgs
    Exit
}

if (-not $RunAsSystemMode -and ([string]::IsNullOrWhiteSpace($GlobalScriptPath) -or -not (Test-Path $GlobalScriptPath))) {
    Write-Host "`n[ABBBRUCH] Das Skript weiss nicht, wo es gespeichert ist (Pfad ist leer)!" -ForegroundColor Red
    Write-Host "Der Hintergrund-Task wuerde sonst ins Leere laufen." -ForegroundColor Yellow
    Write-Host "Loesung: Speichere den Text als .ps1 ab und starte es direkt.`n"
    Exit
}

# --- GUIDS ---
$GUIDs = @{
    CustomScheme   = "ca6ce3ef-da1a-45a5-875e-06940d5890d2"
    ButtonsSub     = "4f971e89-eebd-4455-a8de-9e59040e7347"; LidAction = "5ca83367-6e45-459f-a27b-476b1d01c936"; PowerBtn = "7648efa3-dd9c-4e3e-b566-50f929386280"; SleepBtn = "96996bc0-ad50-47ec-923b-6f41874dd9eb"
    DisplaySub     = "7516b95f-f776-4464-8c53-06167f40cc99"; DisplayOff = "3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e"; LockScreenOff = "facac523-3045-4f54-a919-0c2c866aef81"
    SleepSub       = "238c9fa8-0aad-41ed-83f4-97be242c8f20"; SleepAfter = "29f6c1db-86da-48c5-9fdb-f2b67b1f44da"; HibernAfter = "9d7815a6-7ee4-497e-8888-515a05f02364"
    DiskSub        = "0012ee47-9041-4b5d-9b77-535fba8b1442"; DiskSpindown = "6738e2c4-e8a5-4a42-b16a-e040e769756e"
    CpuSub         = "54533251-82be-4824-96c1-47b60b740d00"; CoolingPolicy = "94d3a615-a899-4ac5-ae2b-e4d8f634367f"; CpuMin = "893dee8e-2bef-41e0-89c6-b55d0929964c"; CpuMax = "bc5038f7-23e0-4960-96da-33abaf5935ec"
    WlanSub        = "19cbb8fa-0579-4dbd-8ee2-434a68d2034f"; WlanWireless = "12bbebe6-58d6-4636-95bb-3217ef867c1a"
    PcieSub        = "50cee402-dd81-4734-99ac-7456ebefd538"; PcieAspm = "ee12f906-d277-404b-b6da-e5fa1579645d"
}

# --- FUNKTIONEN ---
function Set-CoreReg {
    param ([string]$Scheme, [string]$Sub, [string]$Setting, [int]$AC, [int]$DC)
    $regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\$Scheme\$Sub\$Setting"
    if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
    Set-ItemProperty -Path $regPath -Name "ACSettingIndex" -Value $AC -Type DWord -Force | Out-Null
    Set-ItemProperty -Path $regPath -Name "DCSettingIndex" -Value $DC -Type DWord -Force | Out-Null
}

function Get-CurrentPowerState {
    param ([string]$Title)
    Write-Host "`n=== $Title ===" -ForegroundColor Cyan
    $activeScheme = (powercfg /getactivescheme) -match '([0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12})' | Out-Null; $activeScheme = if($matches){$matches[1]}else{"381b4222-f694-41f0-9685-ff5bb260df2e"}
    Write-Host "Aktives Profil GUID: $activeScheme" -ForegroundColor DarkGray
    
    # 1. Helper fuer Energie-Profile
    $readReg = {
        param($sub, $set, $type)
        try { return (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\$activeScheme\$sub\$set" -Name "${type}SettingIndex" -ErrorAction Stop)."${type}SettingIndex" } catch { return $null }
    }

    $formatMin = {
        param($val)
        if ($val -eq $null) { return "N/A" }
        $m = $val / 60
        if ($m -eq 0) { return "0 (Niemals)" }
        return "$m Min"
    }

    $formatAct = {
        param($val)
        if ($val -eq $null) { return "N/A" }
        switch ($val) {
            0 { return "0 (Nichts tun)" }
            1 { return "1 (Standby)" }
            2 { return "2 (Ruhezustand)" }
            3 { return "3 (Herunterfahren)" }
            default { return $val }
        }
    }

    # 2. Helper fuer RDP Settings (Rechnet ms in Min um)
    $readRdp = {
        param($name)
        try { return (Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services" -Name $name -ErrorAction Stop).$name } catch { return $null }
    }
    
    $formatRdpTime = {
        param($val)
        if ($val -eq $null) { return "N/A (Nicht konfiguriert)" }
        if ($val -eq 0) { return "0 (Endlos / Deaktiviert)" }
        return "$([math]::Round($val / 60000, 1)) Min" 
    }
    
    $formatRdpPrompt = {
        param($val)
        if ($val -eq $null) { return "N/A (Nicht konfiguriert)" }
        if ($val -eq 0) { return "0 (Nicht erzwungen)" }
        if ($val -eq 1) { return "1 (Immer abfragen)" }
        return $val
    }

    $results = @()
    
    # TIMEOUTS
    $results += [PSCustomObject]@{ Einstellung = "[SYS] Display Timeout"; AC = &$formatMin(&$readReg $GUIDs.DisplaySub $GUIDs.DisplayOff "AC"); DC = &$formatMin(&$readReg $GUIDs.DisplaySub $GUIDs.DisplayOff "DC") }
    $results += [PSCustomObject]@{ Einstellung = "[SYS] Sperrbildschirm Timeout"; AC = &$formatMin(&$readReg $GUIDs.DisplaySub $GUIDs.LockScreenOff "AC"); DC = &$formatMin(&$readReg $GUIDs.DisplaySub $GUIDs.LockScreenOff "DC") }
    $results += [PSCustomObject]@{ Einstellung = "[SYS] Standby nach"; AC = &$formatMin(&$readReg $GUIDs.SleepSub $GUIDs.SleepAfter "AC"); DC = &$formatMin(&$readReg $GUIDs.SleepSub $GUIDs.SleepAfter "DC") }
    $results += [PSCustomObject]@{ Einstellung = "[SYS] Ruhezustand nach"; AC = &$formatMin(&$readReg $GUIDs.SleepSub $GUIDs.HibernAfter "AC"); DC = &$formatMin(&$readReg $GUIDs.SleepSub $GUIDs.HibernAfter "DC") }
    $results += [PSCustomObject]@{ Einstellung = "[SYS] Festplatte abschalten"; AC = &$formatMin(&$readReg $GUIDs.DiskSub $GUIDs.DiskSpindown "AC"); DC = &$formatMin(&$readReg $GUIDs.DiskSub $GUIDs.DiskSpindown "DC") }

    # SCHALTER
    $results += [PSCustomObject]@{ Einstellung = "[BTN] Aktion: Zuklappen"; AC = &$formatAct(&$readReg $GUIDs.ButtonsSub $GUIDs.LidAction "AC"); DC = &$formatAct(&$readReg $GUIDs.ButtonsSub $GUIDs.LidAction "DC") }
    $results += [PSCustomObject]@{ Einstellung = "[BTN] Aktion: Netztaste"; AC = &$formatAct(&$readReg $GUIDs.ButtonsSub $GUIDs.PowerBtn "AC"); DC = &$formatAct(&$readReg $GUIDs.ButtonsSub $GUIDs.PowerBtn "DC") }
    
    # GRACE PERIOD
    $graceVal = try { (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" -Name "ScreenSaverGracePeriod" -ErrorAction Stop).ScreenSaverGracePeriod } catch { "N/A" }
    $graceStr = if ($graceVal -eq 0) { "0 Sek (Sofort sperren)" } elseif ($graceVal -ne "N/A") { "$graceVal Sek (Schonfrist)" } else { "N/A" }
    $results += [PSCustomObject]@{ Einstellung = "[EXT] Sperrbildschirm Schonfrist"; AC = $graceStr; DC = $graceStr }

    # RDP OVERRIDES (Wenden sich systemweit an, daher AC und DC Spalte identisch)
    $results += [PSCustomObject]@{ Einstellung = "[RDP] Inaktivitaetslimit (MaxIdleTime)"; AC = &$formatRdpTime(&$readRdp "MaxIdleTime"); DC = &$formatRdpTime(&$readRdp "MaxIdleTime") }
    $results += [PSCustomObject]@{ Einstellung = "[RDP] Trennungszeitlimit (MaxDisconnectionTime)"; AC = &$formatRdpTime(&$readRdp "MaxDisconnectionTime"); DC = &$formatRdpTime(&$readRdp "MaxDisconnectionTime") }
    $results += [PSCustomObject]@{ Einstellung = "[RDP] Maximale Laufzeit (MaxConnectionTime)"; AC = &$formatRdpTime(&$readRdp "MaxConnectionTime"); DC = &$formatRdpTime(&$readRdp "MaxConnectionTime") }
    $results += [PSCustomObject]@{ Einstellung = "[RDP] Passwort beim Reconnect"; AC = &$formatRdpPrompt(&$readRdp "fPromptForPassword"); DC = &$formatRdpPrompt(&$readRdp "fPromptForPassword") }

    $results | Format-Table -AutoSize | Out-String | Write-Host
}

# --- SYSTEM-MODUS ---
if ($RunAsSystemMode) {
    if (Get-Service -Name "LenovoSmartStandby" -ErrorAction SilentlyContinue) { Stop-Service -Name "LenovoSmartStandby" -Force -ErrorAction SilentlyContinue }

    $gpoPath = "HKLM:\SOFTWARE\Policies\Microsoft