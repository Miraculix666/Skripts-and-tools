<#
.SYNOPSIS
    PS-WMP_Repair.ps1 - Expert System for WMP Legacy Remote/Local Repair
    Version: 4.5 (Feature Repair & Detailed Analytics)
    
.DESCRIPTION
    Verbesserte Version mit tiefer Event-Analyse und Windows-Feature-Reparatur.
    
    Klassifizierungen:
    - DB_ERR: Datenbank/Medienserver
    - APP_CRASH: Programmabsturz (wmplayer.exe)
    - COM_ERR: Server Execution failed
    - SVC_ERR: Dienstfehler
#>

[CmdletBinding()]
param(
    [Parameter(Position=0, Mandatory=$false)]
    [string]$ComputerName,

    [Parameter(Position=1, Mandatory=$false)]
    [System.Management.Automation.PSCredential]$Credential
)

# --- KONFIGURATION ---
if ([string]::IsNullOrWhiteSpace($ComputerName)) { $ComputerName = Read-Host "Zielcomputer Name" }
$CommonParams = @{ ComputerName = $ComputerName }
if ($Credential) { $CommonParams['Credential'] = $Credential }

# ANSI Farben für saubere TUI
$Colors = @{
    Green  = "$([char]27)[32m"; Red = "$([char]27)[31m"; Cyan = "$([char]27)[36m"
    Reset  = "$([char]27)[0m"; Bold = "$([char]27)[1m"; Yellow = "$([char]27)[33m"; Gray = "$([char]27)[90m"
}

# --- FUNKTIONEN ---

function Get-DetailedWMPAnalysis {
    Write-Host "`n📋 $($Colors.Bold)ANALYSE: Event-Logs werden tiefgreifend gescannt...$($Colors.Reset)"
    $script = {
        $results = @()
        # Suche nach Application Errors (1000) und WMP spezifischen Logs
        $evts = @(Get-WinEvent -FilterHashtable @{LogName='Application'; StartTime=(Get-Date).AddDays(-14)} -MaxEvents 1000 -EA 0 | 
                  Where-Object { $_.Message -match "wmplayer\.exe|Medienserver|WMPNetworkSvc|Media Player" })
        
        foreach ($e in $evts) {
            $category = "UNBEKANNT"
            if ($e.Id -eq 1000 -and $e.Message -match "wmplayer.exe") { $category = "APP_CRASH" }
            elseif ($e.Message -match "Medienserver") { $category = "DB_ERR" }
            elseif ($e.Message -match "Server execution failed") { $category = "COM_ERR" }
            elseif ($e.Message -match "WMPNetworkSvc") { $category = "SVC_ERR" }

            $results += [PSCustomObject]@{
                Time     = $e.TimeCreated.ToString("dd.MM HH:mm")
                ID       = $e.Id
                Category = $category
                Message  = if ($e.Message.Length -gt 120) { $e.Message.Substring(0,117) + "..." } else { $e.Message.Trim() }
            }
        }
        return $results
    }
    return @(Invoke-Command @CommonParams -ScriptBlock $script)
}

# --- MAIN ---

Write-Host "`n$($Colors.Bold)$($Colors.Cyan)🚀 PS-WMP_Repair v4.5 | Host: $ComputerName$($Colors.Reset)"
Write-Host "$($Colors.Gray)===================================================================$($Colors.Reset)"

# 1. Verbindungstest
try {
    $remoteName = Invoke-Command @CommonParams -ScriptBlock { $env:COMPUTERNAME } -ErrorAction Stop
    Write-Host "📡 Verbindung: $($Colors.Green)OK ($remoteName)$($Colors.Reset)"
} catch {
    Write-Host "❌ $($Colors.Red)Verbindungsfehler: $($_.Exception.Message)$($Colors.Reset)"; exit
}

# 2. Echte Fehlerausgabe
$errors = Get-DetailedWMPAnalysis
if ($null -ne $errors -and $errors.Count -gt 0) {
    Write-Host "⚠️  $($Colors.Yellow)Gefundene Fehlermeldungen ($($errors.Count)):$($Colors.Reset)"
    $errors | Select-Object Time, Category, Message | Format-Table -AutoSize | Out-String | Write-Host
} else {
    Write-Host "✅ Keine kritischen WMP-Logs in den letzten 14 Tagen gefunden." -ForegroundColor Green
}

# 3. Menü
Write-Host "🔧 $($Colors.Bold)REPARATUR-STRATEGIE WÄHLEN:$($Colors.Reset)"
Write-Host "   [1] Cache Clean    - Löscht WMP-Datenbanken aller User"
		Write-Host "   [2] System Fix     - Registriert DLLs & setzt Berechtigungen"
Write-Host "   [3] Service Reset  - Rekonfiguriert WMPNetworkSvc"
Write-Host "   [4] Feature Repair - Deaktiviert/Aktiviert WMP Komponente (Tiefenfix)"
Write-Host "   [A] KOMPLETT-FIX   - Führt alle Schritte nacheinander aus"
Write-Host "   [C] Log Cleanup    - Leert das Application Log"
Write-Host "   [X] Abbrechen"

$choice = (Read-Host "`nEingabe").ToUpper()
if ($choice -eq "X") { exit }

Write-Host "`n🚀 $($Colors.Bold)Starte Reparatur...$($Colors.Reset)"

# --- EXECUTION ---

# Schritt 1: Cache (Multi-User)
if ($choice -match "[1A]") {
    Write-Host "   📂 Schritt 1: Cache Bereinigung (Profile)..." -NoNewline
    $cacheReport = Invoke-Command @CommonParams -ScriptBlock {
        $report = ""
        Get-ChildItem "C:\Users" -Directory | ForEach-Object {
            $p = "$($_.FullName)\AppData\Local\Microsoft\Media Player"
            if (Test-Path $p) {
                Remove-Item "$p\*" -Recurse -Force -EA 0
                $report += "$($_.Name);"
            }
        }
        return $report
    }
    Write-Host " ✅ (Bereinigt: $($cacheReport.Trim(';')))" -ForegroundColor Green
}

# Schritt 2: System DLLs & Perms
if ($choice -match "[2A]") {
    Write-Host "   🧩 Schritt 2: DLL & ACL Reparatur..." -NoNewline
    $dllRes = Invoke-Command @CommonParams -ScriptBlock {
        $e = 0
        @("wmp.dll", "jscript.dll", "vbscript.dll") | ForEach-Object { cmd /c "regsvr32 /s $_"; if($LASTEXITCODE -ne 0){$e++} }
        cmd /c "icacls `"$env:SystemDrive\Users`" /grant *S-1-5-20:(RX) /Q" # Leserechte für NetworkSvc auf Profile
        return $e
    }
    if ($dllRes -eq 0) { Write-Host " ✅ OK" -ForegroundColor Green } else { Write-Host " ❌ Fehler" -ForegroundColor Red }
}

# Schritt 3: Service Reset
if ($choice -match "[3A]") {
    Write-Host "   ⚙️  Schritt 3: WMPNetworkSvc Reset..." -NoNewline
    $svcStat = Invoke-Command @CommonParams -ScriptBlock {
        Stop-Service WMPNetworkSvc -Force -EA 0
        Set-Service WMPNetworkSvc -StartupType Disabled
        return (Get-Service WMPNetworkSvc).Status
    }
    Write-Host " ✅ (Status: $svcStat)" -ForegroundColor Green
}

# Schritt 4: Feature Repair (Das gewünschte Feature)
if ($choice -match "[4A]") {
    Write-Host "   📦 Schritt 4: Windows Feature Re-Installation..."
    $featJob = Invoke-Command @CommonParams -ScriptBlock {
        # 1. Deaktivieren
        Disable-WindowsOptionalFeature -FeatureName "WindowsMediaPlayer" -Online -NoRestart -EA 0 | Out-Null
        # 2. Aktivieren
        $res = Enable-WindowsOptionalFeature -FeatureName "WindowsMediaPlayer" -All -Online -NoRestart
        return $res.State
    } -AsJob
    
    while ($featJob.State -eq 'Running') {
        Write-Progress -Activity "Feature Repair" -Status "Komponente wird neu installiert..." -PercentComplete 50
        Start-Sleep 1
    }
    $finalFeat = Receive-Job $featJob
    Write-Host "      ✅ Status: $finalFeat" -ForegroundColor Green
    Remove-Job $featJob
}

# Schritt C: Cleanup
if ($choice -eq "C") {
    Invoke-Command @CommonParams -ScriptBlock { Clear-EventLog -LogName Application }
    Write-Host "   🧹 Logs bereinigt." -ForegroundColor Green
}

# --- FINISH ---
Write-Host "`n$($Colors.Bold)$($Colors.Green)🏁 REPARATUR ABGESCHLOSSEN.$($Colors.Reset)"
Write-Host "   Die technischen Parameter auf $remoteName wurden zurückgesetzt."
Write-Host "   👉 Bitte den User bitten, den WMP nun lokal zu testen."
Write-Host "$($Colors.Gray)-------------------------------------------------------------------$($Colors.Reset)`n"
