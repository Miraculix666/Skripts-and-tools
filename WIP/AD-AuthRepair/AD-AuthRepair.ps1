<#
.SYNOPSIS
    Repariert Authentifizierungsprobleme (Kerberos, Secure Channel) und prüft AD-Zugriff.

.DESCRIPTION
    Dateiname: AD-AuthRepair.ps1
    Version 1.8
    
    https://gemini.google.com/app/7a6c8ab18bc8dc70
    https://gemini.google.com/share/ea42b884d872

    Änderungen in v1.8:
    - FIX: 'Finally'-Block gehärtet: Kurze Wartezeit vor dem Schließen, um versehentliches Beenden zu verhindern.
    - MSG: Klarerer Hinweis am Ende, dass das Fenster offen bleibt.
    
    Änderungen in v1.7:
    - FIX: Robustere Eingabeprüfung (Trim).
    
.PARAMETER Force
    Führt Aktionen ohne Rückfragen aus (Batch-Modus).

.PARAMETER SkipGPUpdate
    Überspringt die Aktualisierung der Gruppenrichtlinien.

.PARAMETER TestIdentity
    Optional: Ein spezifischer Benutzername (SAMAccountName), gegen den die AD-Verbindung geprüft werden soll.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)][switch]$Force,
    [Parameter(Mandatory=$false)][switch]$SkipGPUpdate,
    [Parameter(Mandatory=$false, Position=0)][string]$TestIdentity
)

# ---------------------------------------------------------------------------
# INITIALISIERUNG
# ---------------------------------------------------------------------------
[System.Threading.Thread]::CurrentThread.CurrentCulture = "de-DE"
$Host.UI.RawUI.WindowTitle = "AuthRepair Tool v1.8 - Diagnose & Reparatur"

function Write-Log {
    param([string]$Message, [string]$Level = "Info")
    $timestamp = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
    $color = switch ($Level) { "Info" {"Gray"} "Warning" {"Yellow"} "Error" {"Red"} "Success" {"Green"} }
    Write-Host "[$timestamp] [$($Level.ToUpper())] $Message" -ForegroundColor $color
}

function Assert-Admin {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "Das Skript benötigt Administrator-Rechte! Bitte als Admin neu starten."
    }
}

function Test-ADConnectivity {
    param([string]$TargetUser)
    
    try {
        if (Get-Module -ListAvailable ActiveDirectory) {
            if (-not [string]::IsNullOrWhiteSpace($TargetUser)) {
                # Strikter Test gegen das spezifische Ziel
                Write-Verbose "Prüfe Zugriff auf spezifische Identität: $TargetUser"
                $null = Get-ADPrincipalGroupMembership -Identity $TargetUser -ErrorAction Stop
                return $true
            }
            else {
                # Standard Test (Current User)
                $null = Get-ADUser -Identity $env:USERNAME -Properties Created -ErrorAction Stop
                return $true
            }
        }
        else {
            # Fallback falls kein RSAT installiert ist (nur Verbindungstest)
            $nltest = nltest /sc_query:$env:USERDOMAIN 2>&1
            if ($nltest -match "Success") { return $true } else { throw $nltest }
        }
    }
    catch {
        return $_
    }
}

# ---------------------------------------------------------------------------
# HAUPTPROGRAMM
# ---------------------------------------------------------------------------

try { # Begin Hauptblock für Fehlerfang

    Clear-Host
    Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host " AuthRepair Tool v1.8 - Diagnose & Reparatur " -ForegroundColor Cyan
    Write-Host "------------------------------------------------------------" -ForegroundColor Cyan

    # 1. Admin Prüfung
    Assert-Admin

    # Domäne ermitteln
    try {
        $domainName = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name
        Write-Log "Domäne erkannt: $domainName" "Info"
    }
    catch {
        Write-Log "Konnte Domänennamen nicht automatisch ermitteln." "Warning"
        $domainName = $env:USERDOMAIN
    }

    # Interaktive Abfrage für TestIdentity
    if (-not $TestIdentity -and -not $Force) {
        Write-Host "`nHINWEIS: Wenn Befehle für einen speziellen User (z.B. nwob4746) fehlschlagen," -ForegroundColor Gray
        Write-Host "geben Sie dessen Namen hier ein. Sonst drücken Sie nur ENTER." -ForegroundColor Gray
        $inputIdentity = Read-Host "Zu prüfende Identität (SAMAccountName)"
        if (-not [string]::IsNullOrWhiteSpace($inputIdentity)) {
            $TestIdentity = $inputIdentity
        }
    }

    # SCHRITT 0: DIAGNOSE
    Write-Host "`n[1/5] Initiale Diagnose..." -ForegroundColor Yellow
    if ($TestIdentity) { Write-Log "Prüfe Zugriff auf Ziel: $TestIdentity" "Info" }

    $adStatus = Test-ADConnectivity -TargetUser $TestIdentity
    $needsForceRepair = $false

    if ($adStatus -eq $true) {
        Write-Log "AD-Zugriff war ERFOLGREICH (für geprüfte Identität)." "Success"
        
        if (-not $Force) {
            Write-Host "`nWARNUNG: Manchmal ist die Diagnose 'grün', obwohl versteckte Probleme bestehen." -ForegroundColor Yellow
            $manualForce = Read-Host "Möchten Sie die Reparatur trotzdem ERZWINGEN? (Empfohlen bei Fehler 1244) [J/N]"
            if ($manualForce -match "^(j|y|ja|yes)$") {
                $needsForceRepair = $true
                Write-Log "Reparatur wird manuell erzwungen." "Warning"
            }
        }
    }
    else {
        Write-Log "AD-Zugriff FEHLGESCHLAGEN." "Error"
        Write-Log "Details: $($adStatus.ToString())" "Warning"
        
        if ($adStatus.ToString() -match "1244" -or $adStatus.ToString() -match "authentifiziert" -or $adStatus.ToString() -match "Authenticated") {
            Write-Log "-> Authentifizierungsfehler erkannt. Reparatur wird ERZWINGEN." "Warning"
            $needsForceRepair = $true
        }
    }

    # Bestätigung
    if (-not $Force) {
        Write-Host "`n------------------------------------------------------------"
        Write-Host "Maßnahmen:"
        Write-Host " 1. Kerberos Tickets löschen (User & System)"
        Write-Host " 2. Secure Channel reparieren (Passwort-Reset Computerkonto)"
        Write-Host " 3. Bei Bedarf: Netlogon-Dienst neu starten (vermeidet Reboot)"
        if ($needsForceRepair) { Write-Host " ! Reparatur wird ERZWUNGEN" -ForegroundColor Red }
        Write-Host "------------------------------------------------------------"
        
        $response = Read-Host "Möchten Sie starten? (J/N)"
        # FIX: Leerzeichen trimmen und Variable prüfen
        if ($null -eq $response -or $response.Trim() -notmatch "^(j|y|ja|yes)$") {
            Write-Log "Abbruch durch Benutzer (Eingabe war '$response')." "Warning"
            # Return springt zum Finally-Block -> Fenster bleibt offen
            return 
        }
    }

    # SCHRITT 1: Kerberos Tickets
    Write-Host "`n[2/5] Bereinige Kerberos-Tickets..." -ForegroundColor Yellow
    klist purge | Out-Null
    Write-Log "Benutzer-Tickets bereinigt." "Success"

    try {
        Start-Process -FilePath "klist.exe" -ArgumentList "-li 0x3e7 purge" -NoNewWindow -Wait
        Write-Log "System/Computer-Tickets bereinigt." "Success"
    } catch { Write-Log "System-Tickets konnten nicht bereinigt werden." "Info" }

    # SCHRITT 2: Secure Channel
    Write-Host "`n[3/5] Prüfe und repariere Secure Channel..." -ForegroundColor Yellow
    $isSecure = Test-ComputerSecureChannel -Verbose:$false

    if ($isSecure -and -not $needsForceRepair) {
        Write-Log "Secure Channel ist laut Windows intakt. Keine Aktion." "Success"
    }
    else {
        if ($needsForceRepair) { Write-Log "Erzwinge Reparatur..." "Warning" }
        else { Write-Log "Secure Channel ist unterbrochen." "Warning" }
        
        try {
            $repair = Test-ComputerSecureChannel -Repair -Verbose
            if ($repair) { Write-Log "Reparatur erfolgreich durchgeführt." "Success" }
            else { Write-Log "Reparatur fehlgeschlagen!" "Error" }
        }
        catch { Write-Log "Fehler bei Reparatur: $_" "Error" }
    }

    # SCHRITT 3: GPUpdate
    if (-not $SkipGPUpdate) {
        Write-Host "`n[4/5] Aktualisiere Gruppenrichtlinien..." -ForegroundColor Yellow
        $proc = Start-Process "gpupdate.exe" "/force" -NoNewWindow -Wait -PassThru
        if ($proc.ExitCode -eq 0) { Write-Log "GPUpdate OK." "Success" }
        else { Write-Log "GPUpdate Fehler (Code $($proc.ExitCode))." "Warning" }
    }

    # SCHRITT 4: Abschluss-Verifikation & SOFT RESET LOGIK
    Write-Host "`n[5/5] Abschluss-Verifikation..." -ForegroundColor Yellow
    Start-Sleep -Seconds 2

    $finalCheck = Test-ADConnectivity -TargetUser $TestIdentity

    if ($finalCheck -eq $true) {
        Write-Host "`n------------------------------------------------------------" -ForegroundColor Green
        Write-Log "VERIFIKATION ERFOLGREICH! AD-Zugriff wiederhergestellt." "Success"
        if ($TestIdentity) { Write-Log "Zugriff auf '$TestIdentity' klappt wieder." "Success" }
        Write-Host "------------------------------------------------------------" -ForegroundColor Green
    }
    else {
        Write-Log "Verifikation fehlgeschlagen. Versuche 'Soft-Reset' um Neustart zu vermeiden..." "Warning"
        
        try {
            Write-Host "`n[!] Führe erweiterte Maßnahmen aus..." -ForegroundColor Cyan
            
            # 1. DNS Flush
            Write-Log "Leere DNS Cache (ipconfig /flushdns)..." "Info"
            ipconfig /flushdns | Out-Null
            
            # 2. Netlogon Reset via nltest
            Write-Log "Reset Secure Channel via nltest..." "Info"
            Start-Process "nltest.exe" -ArgumentList "/sc_reset:$domainName" -NoNewWindow -Wait
            
            # 3. Netlogon Service Restart
            Write-Log "Starte Dienst 'Netlogon' neu..." "Info"
            Restart-Service Netlogon -Force -ErrorAction Stop
            
            Write-Log "Warte auf Service-Initialisierung (5 Sek)..." "Info"
            Start-Sleep -Seconds 5
            
            # Letzter Test
            $lastChance = Test-ADConnectivity -TargetUser $TestIdentity
            if ($lastChance -eq $true) {
                 Write-Host "`n------------------------------------------------------------" -ForegroundColor Green
                 Write-Log "SOFT-RESET ERFOLGREICH! Neustart vermieden." "Success"
                 if ($TestIdentity) { Write-Log "Zugriff auf '$TestIdentity' funktioniert." "Success" }
                 Write-Host "------------------------------------------------------------" -ForegroundColor Green
            }
            else {
                throw "Soft-Reset fehlgeschlagen."
            }
        }
        catch {
            Write-Host "`n------------------------------------------------------------" -ForegroundColor Red
            Write-Log "ALLE VERSUCHE FEHLGESCHLAGEN." "Error"
            Write-Host "------------------------------------------------------------" -ForegroundColor Red
            Write-Host " WICHTIG: BITTE STARTEN SIE DEN COMPUTER NEU! " -ForegroundColor Red -BackgroundColor Yellow
            Write-Host " Der LSASS-Cache hält das alte Token fest. Ein Neustart ist zwingend. " -ForegroundColor Red -BackgroundColor Yellow
        }
    }

}
catch {
    Write-Host "`n[FEHLER] Ein unerwarteter Fehler ist aufgetreten:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}
finally {
    Write-Host "`n------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host " ENDE DES SKRIPTS " -ForegroundColor Cyan
    Write-Host " (Bitte prüfen Sie oben stehende Logs, bevor Sie das Fenster schließen)" -ForegroundColor Gray
    
    # Verhindert versehentliches Schließen durch gepufferte Eingaben
    Start-Sleep -Seconds 1
    
    if (-not $Force) {
        Read-Host "Drücken Sie die EINGABETASTE zum Schließen..."
    }
}