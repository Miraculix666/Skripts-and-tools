# PSC_BitLockerRecovery.ps1
# Beschreibung: Kompakter BitLocker-Report mit HTML-Export, Print-Option und Aktiv/Legacy-Hervorhebung.
# Target: PowerShell 5.1 | Windows 10/Server 2019 | RSAT Active Directory
# Localization: DE (German)

<#
.SYNOPSIS
    Sucht BitLocker-Keys, hebt aktive Schlüssel hervor und exportiert optional nach HTML/Druck.
.PARAMETER ComputerName
    Name oder Teilstring des Ziel-Computers.
.PARAMETER Detailed
    Zeigt zusätzlich den DistinguishedName an.
.PARAMETER RemoteFix
    Aktiviert den Remote-Fix-Batchmodus (Suspend/Resume).
.PARAMETER Print
    Generiert den HTML-Bericht und startet direkt den Druck-Dialog.
.PARAMETER HtmlPath
    Eigener Speicherpfad für die HTML-Datei (Standard: .\PSC_BitLockerReport.html).
#>

[CmdletBinding()]
Param(
    [Parameter(Position = 0, Mandatory = $false)]
    [String]$ComputerName,

    [Parameter()]
    [Switch]$Detailed,

    [Parameter()]
    [Switch]$RemoteFix,

    [Parameter()]
    [Switch]$Print,

    [Parameter()]
    [String]$HtmlPath = ".\PSC_BitLockerReport.html"
)

# --- [ INITIALIZATION ] ---
$ErrorActionPreference = "Stop"
$VerbosePreference = "Continue"
$DateFormat = "dd.MM.yyyy HH:mm"

# --- [ PREREQUISITES ] ---
if (-not (Get-Module -ListAvailable ActiveDirectory)) {
    try { Import-Module ActiveDirectory -ErrorAction Stop } catch { Write-Error "RSAT AD Tools benötigt."; return }
}

# --- [ INPUT ] ---
if ([string]::IsNullOrWhiteSpace($ComputerName)) {
    Write-Host "--- BITLOCKER SUCHE ---" -ForegroundColor Cyan
    $ComputerName = Read-Host "Computername oder Teilstring"
}

$SearchFilter = if ($ComputerName.Contains("*")) { $ComputerName } else { "*$ComputerName*" }

# --- [ DATA RETRIEVAL ] ---
$Computers = Get-ADComputer -Filter "Name -like '$SearchFilter'" -Properties DistinguishedName | Sort-Object Name

if (-not $Computers) {
    Write-Warning "Keine Treffer für '$SearchFilter' gefunden."
    return
}

# --- [ HTML HEAD PREPARATION ] ---
$HtmlHeader = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <title>BitLocker Recovery Report</title>
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; margin: 30px; color: #333; line-height: 1.4; }
        h1 { color: #005a9c; border-bottom: 2px solid #005a9c; padding-bottom: 10px; }
        .computer-block { margin-bottom: 35px; padding: 15px; border: 1px solid #ddd; background: #fafafa; border-radius: 4px; page-break-inside: avoid; }
        .computer-name { font-size: 1.3em; font-weight: bold; color: #111; margin: 0 0 5px 0; }
        .dn { font-size: 0.85em; color: #666; margin-bottom: 10px; }
        .key-box { margin: 5px 0; padding: 8px 12px; border-radius: 3px; font-family: Consolas, monospace; }
        .active { background-color: #d4edda; color: #155724; border-left: 5px solid #28a745; font-weight: bold; }
        .legacy { background-color: #e2e3e5; color: #383d41; border-left: 5px solid #6c757d; font-size: 0.95em; opacity: 0.8; }
        .meta { font-size: 0.8em; color: #555; display: inline-block; margin-left: 15px; }
        @media print { body { margin: 0; } .computer-block { page-break-inside: avoid; } }
    </style>
</head>
<body>
    <h1>BitLocker Wiederherstellungsschlüssel Bericht</h1>
    <p>Generiert am: $(Get-Date -Format $DateFormat)</p>
"@

$HtmlBody = ""

# --- [ CONSOLE & HTML OUTPUT GENERATION ] ---
Write-Host "`nErgebnis: $($Computers.Count) Objekte gefunden." -ForegroundColor Gray
Write-Host ("==============================================================")

foreach ($Comp in $Computers) {
    # Konsole Kopfzeile (Sauber ohne -Bold String-Artefakte)
    Write-Host "$($Comp.Name)" -ForegroundColor White -Bold

    # HTML Block starten
    $HtmlBody += "<div class='computer-block'>"
    $HtmlBody += "  <div class='computer-name'>$($Comp.Name)</div>"

    if ($Detailed) {
        Write-Host "  DN: $($Comp.DistinguishedName)" -ForegroundColor Gray
        $HtmlBody += "  <div class='dn'>DN: $($Comp.DistinguishedName)</div>"
    }

    # Recovery Keys abrufen (Neueste zuerst)
    $Keys = Get-ADObject -Filter "objectClass -eq 'msFVE-RecoveryInformation'" `
        -SearchBase $Comp.DistinguishedName `
        -Properties msFVE-RecoveryPassword, whenCreated, Name | Sort-Object whenCreated -Descending

    if ($Keys) {
        $isFirst = $true
        foreach ($K in $Keys) {
            $DeviceId = if ($K.Name -match "\{(.*)\}") { "{" + $Matches[1] + "}" } else { $K.Name }
            $DateStr = $K.whenCreated.ToString($DateFormat)

            if ($isFirst) {
                # --- AKTIVER SCHLÜSSEL ---
                Write-Host "  ID: $DeviceId" -ForegroundColor DarkGray
                Write-Host "  [AKTIV] KEY: " -NoNewline -ForegroundColor Green -Bold
                Write-Host "$($K.'msFVE-RecoveryPassword')" -ForegroundColor White -BackgroundColor DarkGreen
                if ($Detailed) { Write-Host "  TAG: $DateStr" -ForegroundColor Gray }

                $HtmlBody += @"
                <div class='key-box active'>
                    ID: $DeviceId &nbsp;&nbsp;|&nbsp;&nbsp; KEY: $($K.'msFVE-RecoveryPassword')
                    <span class='meta'>[AKTIV - Erstellt: $DateStr]</span>
                </div>
"@
                $isFirst = $false
            }
            else {
                # --- LEGACY SCHLÜSSEL ---
                Write-Host "  ID: $DeviceId" -ForegroundColor DarkGray
                Write-Host "  [LEGACY] KEY: " -NoNewline -ForegroundColor Gray
                Write-Host "$($K.'msFVE-RecoveryPassword')" -ForegroundColor DarkGray
                if ($Detailed) { Write-Host "  TAG: $DateStr" -ForegroundColor DarkGray }

                $HtmlBody += @"
                <div class='key-box legacy'>
                    ID: $DeviceId &nbsp;&nbsp;|&nbsp;&nbsp; KEY: $($K.'msFVE-RecoveryPassword')
                    <span class='meta'>[LEGACY - Erstellt: $DateStr]</span>
                </div>
"@
            }
        }
    }
    else {
        Write-Host "  [!] Keine BitLocker-Daten gefunden." -ForegroundColor Red
        $HtmlBody += "<div class='key-box' style='color:red;'>[!] Keine BitLocker-Daten im AD hinterlegt.</div>"
    }

    $HtmlBody += "</div>"

    # Zwei Absätze Abstand in der Konsole
    Write-Host "`n"
    Write-Host ("-" * 62)
    Write-Host "`n"
}

# HTML abschließen und speichern
$HtmlFooter = "</body></html>"
$FinalHtml = $HtmlHeader + $HtmlBody + $HtmlFooter
$FinalHtml | Out-File -FilePath $HtmlPath -Encoding utf8

Write-Host "--- BERICHT ABGESCHLOSSEN ---" -ForegroundColor Cyan
Write-Host "HTML-Bericht gespeichert unter: (Get-Item $HtmlPath).FullName" -ForegroundColor Gray

# --- [ AUTOMATIC PRINT TRIGGER ] ---
if ($Print) {
    Write-Verbose "[PRINT] Starte Druckprozess für die HTML-Datei..."
    try {
        # Nutzt das Standard-Druck-Verb des Systems für HTML-Dateien
        Start-Process -FilePath $HtmlPath -Verb Print -WindowStyle Hidden
        Write-Host "[SUCCESS] Druckauftrag wurde an den Standarddrucker gesendet." -ForegroundColor Green
    }
    catch {
        Write-Warning "Drucken fehlgeschlagen. Bitte öffnen Sie die Datei manuell zum Drucken: $HtmlPath"
    }
}

# --- [ REMOTE FIX FUNKTION (BATCH) ] ---
$RunFix = $RemoteFix
if (-not $RunFix -and (Test-Connection -ComputerName "localhost" -Count 1 -Quiet)) {
    # Verhindert Blockade im non-interactive
    Write-Host "`nMöchten Sie bei den gefundenen PCs BitLocker remote reparieren (Fix Boot-Loop)?" -ForegroundColor Cyan
    $StartInteraktiv = Read-Host "Ausführen? (J/N)"
    if ($StartInteraktiv -match '^[Jj]') { $RunFix = $true }
}

if ($RunFix) {
    Write-Host "`n--- STARTE REMOTE FIX (BATCH-MODUS) ---" -ForegroundColor Cyan
    $ConfirmAll = Read-Host "Soll der BitLocker-Reset für ALLE $($Computers.Count) Clients ausgeführt werden? (J/N)"
    if ($ConfirmAll -match '^[Jj]') {
        foreach ($Comp in $Computers) {
            Write-Host "`nVerarbeite Client: $($Comp.Name)" -ForegroundColor Cyan
            if (Test-Connection -ComputerName $Comp.Name -Count 1 -Quiet) {
                try {
                    Invoke-Command -ComputerName $Comp.Name -ScriptBlock {
                        Suspend-BitLocker -MountPoint "C:" -RebootCount 1 -ErrorAction Stop
                    }
                    Write-Host "  [OK] BitLocker pausiert. Warte 5 Sekunden..." -ForegroundColor Green
                    Start-Sleep -Seconds 5
                    Invoke-Command -ComputerName $Comp.Name -ScriptBlock {
                        Resume-BitLocker -MountPoint "C:" -ErrorAction Stop
                    }
                    Write-Host "  [SUCCESS] BitLocker wurde reaktiviert!" -ForegroundColor Green
                }
                catch {
                    Write-Error "  [FEHLER] Remote-Befehl fehlgeschlagen: $($_.Exception.Message)"
                }
            }
            else {
                Write-Warning "  [!] Client $($Comp.Name) ist offline."
            }
        }
    }
}

#
