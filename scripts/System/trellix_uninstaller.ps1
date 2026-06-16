<#
.SYNOPSIS
    Automatisiertes Skript zur vollstaendigen Deinstallation von Trellix Endpoint Security (ENS) und dem Trellix Agent.

.DESCRIPTION
    Dieses Skript deinstalliert die Trellix/McAfee ENS-Module in der zwingend vorgeschriebenen Reihenfolge 
    über die Registry-Uninstall-Strings (ohne WMI). Zum Schluss wird der Trellix Agent entfernt.
    Es erfordert zwingend lokale Administratorrechte.

.NOTES
    Autor: Systemadministrator
    Datum: Juni 2026
#>

# -------------------------------------------------------------------------
# 1. Pruefung auf Administratorrechte
# -------------------------------------------------------------------------
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Warning "Dieses Skript benoetigt lokale Administratorrechte!"
    Write-Host "Versuche, das Skript mit erhoehten Rechten neu zu starten..." -ForegroundColor Yellow
    
    try {
        # Startet das Skript neu mit Admin-Rechten
        Start-Process powershell.exe -ArgumentList ("-NoProfile -ExecutionPolicy Bypass -File `"{0}`"" -f $PSCommandPath) -Verb RunAs
        exit
    } catch {
        Write-Error "Fehler beim Anfordern der Administratorrechte. Bitte starten Sie die PowerShell manuell als Administrator."
        exit
    }
}

Write-Host "Administratorrechte bestätigt. Starte Deinstallationsprozess..." -ForegroundColor Green
Write-Host "--------------------------------------------------------"

# -------------------------------------------------------------------------
# 2. Variablen & Hilfsfunktionen
# -------------------------------------------------------------------------

# Basispfad in der Registry fuer Trellix/McAfee Module
$regPath = "HKLM:\SOFTWARE\WOW6432Node\Network Associates\ePolicy Orchestrator\Application Plugins"

# Hilfsfunktion zum sicheren Ausfuehren und Warten
function Invoke-UninstallCommand {
    param (
        [string]$ModuleName,
        [string]$RegKeyName
    )

    $fullPath = Join-Path -Path $regPath -ChildPath $RegKeyName
    $regKey = Get-ItemProperty -Path $fullPath -ErrorAction SilentlyContinue

    if ($null -ne $regKey -and $regKey.'Uninstall Command') {
        $uninstallCmd = $regKey.'Uninstall Command'
        Write-Host "-> Deinstalliere $ModuleName..." -ForegroundColor Cyan
        Write-Host "   Befehl: $uninstallCmd" -ForegroundColor DarkGray
        
        # Nutzen von cmd /c und Start-Process -Wait verhindert, dass MSI-Prozesse sich ueberschneiden
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$uninstallCmd`"" -Wait -NoNewWindow
        Write-Host "   [OK] $ModuleName Deinstallationskommando abgeschlossen." -ForegroundColor Green
    } else {
        Write-Host "-> $ModuleName ist nicht installiert (Schluessel nicht gefunden). Ueberspringe..." -ForegroundColor DarkGray
    }
}

# -------------------------------------------------------------------------
# 3. Deinstallation der ENS-Zusatzmodule (Reihenfolge extrem wichtig!)
# -------------------------------------------------------------------------

# Schritt 3.1: Threat Prevention (AM) deinstallieren
Invoke-UninstallCommand -ModuleName "Threat Prevention (AM)" -RegKeyName "ENDP_AM_1050"

# Schritt 3.2: Advanced Threat Protection (ATP) deinstallieren
Invoke-UninstallCommand -ModuleName "Advanced Threat Protection (ATP)" -RegKeyName "ENDP_ATP_1050"

# Schritt 3.3: Data Exchange Layer (DXL) Client deinstallieren
Invoke-UninstallCommand -ModuleName "Data Exchange Layer Client (DXL)" -RegKeyName "DXL____1000"

# -------------------------------------------------------------------------
# 4. Deinstallation der Endpoint Security Plattform (MUSS ALS LETZTES ENS-MODUL ERFOLGEN)
# -------------------------------------------------------------------------

# Schritt 4.1: Endpoint Security Platform (GS) deinstallieren
Invoke-UninstallCommand -ModuleName "Endpoint Security Plattform (GS)" -RegKeyName "ENDP_GS_1050"

Write-Host "--------------------------------------------------------"
Write-Host "Warte 5 Sekunden, damit Systemprozesse aufraeumen koennen..." -ForegroundColor DarkGray
Start-Sleep -Seconds 5

# -------------------------------------------------------------------------
# 5. Deinstallation des Trellix Agents
# -------------------------------------------------------------------------
Write-Host "Pruefe auf Trellix Agent..." -ForegroundColor Cyan

$agent32 = "C:\Program Files\McAfee\Agent\x86\FrmInst.exe"
$agent64 = "C:\Program Files (x86)\McAfee\Agent\x86\FrmInst.exe"
$agentPath = $null

if (Test-Path $agent64) {
    $agentPath = $agent64
} elseif (Test-Path $agent32) {
    $agentPath = $agent32
}

if ($null -ne $agentPath) {
    Write-Host "-> Deinstalliere Trellix Agent..." -ForegroundColor Cyan
    Write-Host "   Pfad: $agentPath" -ForegroundColor DarkGray
    
    # Ausfuehren mit Parameter /remove=Agent und Warten auf Beendigung
    Start-Process -FilePath $agentPath -ArgumentList "/remove=Agent" -Wait -NoNewWindow
    
    Write-Host "   [OK] Trellix Agent Deinstallation abgeschlossen." -ForegroundColor Green
} else {
    Write-Host "-> Trellix Agent Deinstallationsdatei (FrmInst.exe) nicht gefunden. Vermutlich bereits deinstalliert." -ForegroundColor DarkGray
}

# -------------------------------------------------------------------------
# Abschluss
# -------------------------------------------------------------------------
Write-Host "--------------------------------------------------------"
Write-Host "Deinstallationsprozess vollstaendig durchlaufen." -ForegroundColor Green
Write-Host "Hinweis: Der Rechner sollte nun im ePO-Server bereinigt (geloescht) werden, um die Compliance-Berichte nicht zu verfaelschen." -ForegroundColor Yellow