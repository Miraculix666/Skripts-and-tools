# WinConfig_Hardening_DE.ps1 - Ultimate Tteck-Style Windows Optimization Suite V6.3
<#
.SYNOPSIS
    WinConfig_Hardening_DE.ps1 - Professionelle Windows Härtung & Optimierung (Tteck-Style)

.DESCRIPTION
    Updates in V6.3:
    - Family Safety: Option zum Zulassen von Microsoft-Konten und Kindersicherung (WpcSvc) hinzugefügt.
    - Login: Lokale Benutzerliste bleibt weiterhin sichtbar (aus V6.2).
    - FeedbackHub: Timeout-Fix und Prozess-Stopp (aus V6.1).

.NOTES
    Version: 6.3.0
    PowerShell: 5.1+
#>

[CmdletBinding()]
Param (
    [Parameter(HelpMessage="Batch-Modus: Führt gewählte Aktionen ohne Menü aus.")]
    [switch]$BatchMode
)

# ---------------------------------------------------------------------------
# 1. STYLE & KONFIGURATION
# ---------------------------------------------------------------------------
$ErrorActionPreference = "SilentlyContinue"
$Script:LogFile = "$env:TEMP\WinConfig_V6_3_$(Get-Date -Format 'yyyyMMdd_HHmm').log"

$Theme = @{
    Border = "Magenta"; Text = "White"; Highlight = "Cyan"
    Success = "Green"; Warning = "Yellow"; Error = "Red"; Dim = "DarkGray"
}

# Unicode Symbole
$Tick = [char]0x2714; $Box = [char]0x2610; $Tree = [char]0x251C; $Last = [char]0x2514; $Line = [char]0x2500; $Vert = [char]0x2502

$Global:Cfg = [ordered]@{
    "1_Privacy"     = @{ Active = $false; Name = "🛡️  BSI PRIVACY"; Desc = "Telemetrie, Cortana & Bing-Suche stoppen" }
    "2_Security"    = @{ Active = $false; Name = "🔒  SECURITY & LOGIN"; Desc = "UAC-Hardening & Benutzerliste am Login anzeigen" }
    "3_Remote"      = @{ Active = $false; Name = "🌐  REMOTE ACCESS"; Desc = "RDP (NLA) & WinRM für Admins aktivieren" }
    "4_UIWin11"     = @{ Active = $false; Name = "🎨  MODERN UI"; Desc = "Klassik-Kontextmenü & Taskbar Links" }
    "5_Performance" = @{ Active = $false; Name = "🚀  PERFORMANCE"; Desc = "Ultimate Plan, Hibernate EIN, FastBoot AUS" }
    "6_Debloat"     = @{ Active = $false; Name = "🧹  DEBLOAT"; Desc = "Werbe-Apps & FeedbackHub (Clean-Up)" }
    "7_System"      = @{ Active = $false; Name = "⚙️  SYSTEM TWEAKS"; Desc = "Explorer-Ansichten & Dateiendungen" }
    "8_Family"      = @{ Active = $false; Name = "👨‍👩‍👧‍👦  FAMILY SAFETY"; Desc = "MS-Accounts & Kindersicherung zulassen" }
}

# ---------------------------------------------------------------------------
# 2. CORE FUNKTIONEN
# ---------------------------------------------------------------------------

function Write-Tteck {
    param([string]$Msg, [string]$Type = "INFO")
    $Time = Get-Date -Format "HH:mm:ss"
    switch ($Type) {
        "STEP"  { Write-Host "`n$Line$Line$Line $Msg $Line$Line$Line" -ForegroundColor $Theme.Highlight }
        "OK"    { Write-Host "  $Tick  $Msg" -ForegroundColor $Theme.Success }
        "WARN"  { Write-Host "  !  $Msg" -ForegroundColor $Theme.Warning }
        "ERROR" { Write-Host "  X  $Msg" -ForegroundColor $Theme.Error }
        "INFO"  { Write-Host "  i  $Msg" -ForegroundColor $Theme.Dim }
    }
    Add-Content -Path $Script:LogFile -Value "[$Time] [$Type] $Msg"
}

function Set-RegKey {
    param ([string]$Path, [string]$Name, $Value, [string]$Type = "DWord")
    if (-not (Test-Path $Path)) { New-Item -Path $Path -Force | Out-Null }
    Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type -Force | Out-Null
}

# ---------------------------------------------------------------------------
# 3. TASK MODULE
# ---------------------------------------------------------------------------

function Exec-Privacy {
    Write-Tteck "Konfiguriere BSI-Datenschutz..." "STEP"
    Set-RegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" "AllowTelemetry" 0
    Set-RegKey "HKLM:\SOFTWARE\Microsoft\Windows\Windows Search" "AllowCortana" 0
    Set-RegKey "HKLM:\SOFTWARE\Microsoft\Windows\Windows Search" "DisableWebSearch" 1
    Write-Tteck "Privatsphäre-Einstellungen gesetzt" "OK"
}

function Exec-Security {
    Write-Tteck "Härte System-Sicherheit & Login-Anzeige..." "STEP"
    
    Set-RegKey "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" "DisableCAD" 0
    
    $UAC = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
    Set-RegKey $UAC "ConsentPromptBehaviorAdmin" 2
    
    # MS Account Block: Nur wenn Family Safety NICHT gewünscht ist
    if (-not $Global:Cfg["8_Family"].Active) {
        Set-RegKey $UAC "NoConnectedUser" 3
        Write-Tteck "Microsoft-Konten für Login blockiert (Hardening)" "INFO"
    } else {
        Set-RegKey $UAC "NoConnectedUser" 0
        Write-Tteck "Microsoft-Konten erlaubt (für Family Safety)" "OK"
    }
    
    Set-RegKey $UAC "dontdisplaylastusername" 0
    Set-RegKey $UAC "EnumerateLocalUsersOnDomainJoinedComputers" 1
    
    Write-Tteck "Sicherheits-Parameter & Benutzerliste konfiguriert" "OK"
}

function Exec-FamilySafety {
    Write-Tteck "Konfiguriere Family Safety & MS-Accounts..." "STEP"
    
    # Dienst für Kindersicherung aktivieren (WpcSvc)
    $Svc = Get-Service -Name "WpcSvc" -ErrorAction SilentlyContinue
    if ($Svc) {
        Set-Service -Name "WpcSvc" -StartupType Manual
        Write-Tteck "Dienst für Kindersicherung (WpcSvc) auf Manuell gesetzt" "OK"
    }
    
    # Sicherstellen, dass die Live-ID Anmeldung funktioniert
    Set-RegKey "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" "NoConnectedUser" 0
    Write-Tteck "Microsoft-Konto-Integration ist nun aktiv" "OK"
}

function Exec-Remote {
    Write-Tteck "Aktiviere Fernwartung..." "STEP"
    Set-RegKey "HKLM:\System\CurrentControlSet\Control\Terminal Server" "fDenyTSConnections" 0
    Enable-NetFirewallRule -DisplayGroup "Remote Desktop" -ErrorAction SilentlyContinue
    Enable-PSRemoting -Force -SkipNetworkProfileCheck
    Write-Tteck "RDP & WinRM konfiguriert" "OK"
}

function Exec-Performance {
    Write-Tteck "Optimiere Energie & Start..." "STEP"
    & powercfg -duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61 | Out-Null
    & powercfg /hibernate on
    Set-RegKey "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power" "HiberbootEnabled" 0
    Write-Tteck "Ultimate Power Plan & Hibernate aktiv" "OK"
}

function Exec-Debloat {
    Write-Tteck "Starte App-Bereinigung (Debloat)..." "STEP"
    
    $Apps = @(
        "Microsoft.BingWeather", 
        "Microsoft.WindowsFeedbackHub", 
        "Microsoft.ZuneMusic", 
        "Microsoft.ZuneVideo", 
        "Microsoft.SolitaireCollection",
        "Microsoft.People"
    )

    foreach ($AppName in $Apps) {
        $Package = Get-AppxPackage -Name $AppName -AllUsers
        if ($Package) {
            $FriendlyName = $AppName -replace "Microsoft.ZuneMusic","Groove Musik" -replace "Microsoft.ZuneVideo","Filme & TV"
            Write-Tteck "Verarbeite $FriendlyName..." "INFO"
            
            if ($AppName -like "*FeedbackHub*") {
                Stop-Process -Name "FeedbackHub" -Force -ErrorAction SilentlyContinue
            }

            $Task = Start-Job -ScriptBlock { param($Name); Get-AppxPackage -Name $Name -AllUsers | Remove-AppxPackage } -ArgumentList $AppName
            if (Wait-Job $Task -Timeout 20) {
                Receive-Job $Task | Out-Null
                Write-Tteck "$FriendlyName entfernt" "OK"
            } else {
                Stop-Job $Task
                Write-Tteck "Timeout bei $FriendlyName (Übersprungen)" "WARN"
            }
            Remove-Job $Task
        }
    }
}

function Exec-UI {
    Write-Tteck "Anpassung der UI..." "STEP"
    $CLSID = "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32"
    if (-not (Test-Path $CLSID)) { New-Item -Path $CLSID -Value "" -Force | Out-Null }
    Set-RegKey "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "TaskbarAl" 0
    Write-Tteck "Windows 11 UI Tweaks angewendet" "OK"
}

function Exec-System {
    Write-Tteck "Explorer Tweaks..." "STEP"
    Set-RegKey "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "HideFileExt" 0
    Set-RegKey "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "LaunchTo" 1
    Write-Tteck "Explorer-Ansicht optimiert" "OK"
}

# ---------------------------------------------------------------------------
# 4. TTECK MENÜ
# ---------------------------------------------------------------------------

function Show-ModernMenu {
    while ($true) {
        Clear-Host
        Write-Host "┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓" -ForegroundColor $Theme.Border
        Write-Host "┃          WIN-CONFIG ULTIMATE V6.3 - TTECK STYLE                ┃" -ForegroundColor $Theme.Highlight
        Write-Host "┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛" -ForegroundColor $Theme.Border
        
        $i = 1
        foreach ($Key in $Global:Cfg.Keys) {
            $Item = $Global:Cfg[$Key]
            $Status = if ($Item.Active) { "$Tick" } else { " " }
            $SCol = if ($Item.Active) { $Theme.Success } else { "White" }
            
            Write-Host "  $i. [$Status] " -NoNewline -ForegroundColor $SCol
            Write-Host "$($Item.Name)" -ForegroundColor "White"
            Write-Host "     $Vert   $Line $($Item.Desc)" -ForegroundColor $Theme.Dim
            $i++
        }
        
        Write-Host "`n  [A] Alles (Härtung)   [N] Auswahl Null   [S] STARTEN" -ForegroundColor $Theme.Highlight
        Write-Host "  [Q] BEENDEN" -ForegroundColor $Theme.Error
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor $Theme.Border
        
        $Choice = Read-Host "  Option wählen"
        
        switch ($Choice.ToUpper()) {
            "1" { $Global:Cfg["1_Privacy"].Active = -not $Global:Cfg["1_Privacy"].Active }
            "2" { $Global:Cfg["2_Security"].Active = -not $Global:Cfg["2_Security"].Active }
            "3" { $Global:Cfg["3_Remote"].Active = -not $Global:Cfg["3_Remote"].Active }
            "4" { $Global:Cfg["4_UIWin11"].Active = -not $Global:Cfg["4_UIWin11"].Active }
            "5" { $Global:Cfg["5_Performance"].Active = -not $Global:Cfg["5_Performance"].Active }
            "6" { $Global:Cfg["6_Debloat"].Active = -not $Global:Cfg["6_Debloat"].Active }
            "7" { $Global:Cfg["7_System"].Active = -not $Global:Cfg["7_System"].Active }
            "8" { $Global:Cfg["8_Family"].Active = -not $Global:Cfg["8_Family"].Active }
            "A" { $Global:Cfg.Keys | ForEach-Object { if ($_ -ne "3_Remote" -and $_ -ne "8_Family") { $Global:Cfg[$_].Active = $true } } }
            "N" { $Global:Cfg.Keys | ForEach-Object { $Global:Cfg[$_].Active = $false } }
            "S" { return }
            "Q" { exit }
        }
    }
}

# ---------------------------------------------------------------------------
# 5. START & KEEP-ALIVE
# ---------------------------------------------------------------------------

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "`n[!] BITTE ALS ADMINISTRATOR STARTEN [!]`n" -ForegroundColor $Theme.Error; Read-Host "Taste drücken zum Beenden"; exit
}

# Endlosschleife für das Terminal
while ($true) {
    if (-not $BatchMode) { Show-ModernMenu }

    Clear-Host
    Write-Tteck "Starte gewählte Module..." "STEP"

    if ($Global:Cfg["1_Privacy"].Active)     { Exec-Privacy }
    if ($Global:Cfg["2_Security"].Active)    { Exec-Security }
    if ($Global:Cfg["3_Remote"].Active)      { Exec-Remote }
    if ($Global:Cfg["4_UIWin11"].Active)     { Exec-UI }
    if ($Global:Cfg["5_Performance"].Active) { Exec-Performance }
    if ($Global:Cfg["6_Debloat"].Active)     { Exec-Debloat }
    if ($Global:Cfg["7_System"].Active)      { Exec-System }
    if ($Global:Cfg["8_Family"].Active)      { Exec-FamilySafety }

    Write-Host "`n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor $Theme.Border
    Write-Tteck "Vorgang abgeschlossen!" "OK"
    Write-Tteck "Log: $Script:LogFile" "INFO"
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor $Theme.Border
    
    if ($BatchMode) { break }
    
    Write-Host "`n  [M] Zurück zum Menü    [Q] Beenden" -ForegroundColor $Theme.Highlight
    $Final = Read-Host "  Wahl"
    if ($Final.ToUpper() -eq "Q") { break }
}

Write-Host "`nSkript beendet. Das Fenster bleibt offen. Schließen Sie es manuell oder drücken Sie Enter."
Read-Host