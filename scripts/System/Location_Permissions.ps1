# PS_LocFix_Manager.ps1
# Description: Advanced Location Service Manager & WLAN Fix


# ==============================================================================
# 1. INITIALISIERUNG & ADMIN-CHECK
# ==============================================================================
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red
    Write-Host "BITTE ALS ADMINISTRATOR AUSFUEHREN!" -ForegroundColor Red
    Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red
    pause; exit
}

$LogFile = Join-Path $PSScriptRoot "PS_LocFix_Execution.log"

Function Write-Log {
    Param(
        [Parameter(Mandatory=$true)] [string]$Message, 
        [Parameter(Mandatory=$false)] [string]$Type = "INFO"
    )
    $Stamp = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
    $LogLine = "[$Stamp] [$Type] $Message"
    try {
        Add-Content -Path $LogFile -Value $LogLine -ErrorAction SilentlyContinue
    } catch {}
    
    $Color = switch($Type) { "ERROR" {"Red"} "WARN" {"Yellow"} default {"Gray"} }
    Write-Host $LogLine -ForegroundColor $Color
}

# ==============================================================================
# 2. BENUTZERAUSWAHL
# ==============================================================================
Clear-Host
Write-Host "--- Windows Ortungsdienst-Manager V1.4 ---" -ForegroundColor Cyan
Write-Host "1) Ortungsdienste AKTIVIEREN & WLAN-Scan fixen" -ForegroundColor Green
Write-Host "2) Ortungsdienste DEAKTIVIEREN" -ForegroundColor Red
Write-Host ""
$choice = Read-Host "Bitte waehlen Sie eine Option (1 oder 2, Standard ist 1)"

if ($choice -eq "2") {
    $action = "Disable"
    $regValue = "Deny"
    $privacyValue = 0
    Write-Log "Modus: Deaktivierung gestartet..."
} else {
    $action = "Enable"
    $regValue = "Allow"
    $privacyValue = 1
    Write-Log "Modus: Aktivierung & WLAN-Fix gestartet..."
}

# ==============================================================================
# 3. REGISTRY GUI-FIX (ENTSPERREN DER EINSTELLUNGEN)
# ==============================================================================
Write-Log "Bereinige Registry-Sperren für GUI..."

$Paths = @(
    "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy",
    "HKCU:\Software\Policies\Microsoft\Windows\AppPrivacy",
    "HKLM:\SOFTWARE\Policies\Microsoft\Windows\LocationAndSensors",
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location"
)

foreach ($path in $Paths) {
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
    
    # Entferne blockierende Policy-Einträge
    Remove-ItemProperty -Path $path -Name "LetAppsAccessLocation" -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path $path -Name "DisableLocation" -ErrorAction SilentlyContinue
    
    # Setze expliziten Zugriffswert
    Set-ItemProperty -Path $path -Name "Value" -Value $regValue -Force -ErrorAction SilentlyContinue
}

# Privacy Flag für User-Profil (Apps Zugriff gewähren)
$UserPrivacyPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Privacy"
if (-not (Test-Path $UserPrivacyPath)) { New-Object -TypeName PSObject | Out-Null } # Dummy check
if (-not (Test-Path $UserPrivacyPath)) { New-Item -Path $UserPrivacyPath -Force | Out-Null }
Set-ItemProperty -Path $UserPrivacyPath -Name "LetAppsAccessLocation" -Value $privacyValue -ErrorAction SilentlyContinue

# ==============================================================================
# 4. DIENSTE-STEUERUNG
# ==============================================================================
$Services = @("lfsvc", "RmSvc", "SensorService", "SensrSvc")

foreach ($svcName in $Services) {
    if (Get-Service -Name $svcName -ErrorAction SilentlyContinue) {
        try {
            if ($action -eq "Enable") {
                Set-Service -Name $svcName -StartupType Automatic
                Start-Service -Name $svcName -ErrorAction SilentlyContinue
                Write-Log "Dienst $svcName wurde auf 'Automatisch' gesetzt und gestartet."
            } else {
                Stop-Service -Name $svcName -Force -ErrorAction SilentlyContinue
                Set-Service -Name $svcName -StartupType Disabled
                Write-Log "Dienst $svcName wurde gestoppt und deaktiviert."
            }
        } catch {
            Write-Log "Fehler bei Dienst $svcName : ${ _ }" "ERROR"
        }
    }
}

# ==============================================================================
# 5. WLAN-SCAN FIX (NUR BEI AKTIVIERUNG)
# ==============================================================================
if ($action -eq "Enable") {
    Write-Log "Optimiere WLAN-Scan Parameter..."
    
    # WLAN-Dienst sicherstellen
    Set-Service -Name "WlanSvc" -StartupType Automatic
    Start-Service -Name "WlanSvc" -ErrorAction SilentlyContinue

    # OEM Scan Fix
    $WlanParamPath = "HKLM:\SYSTEM\CurrentControlSet\Services\WlanSvc\Parameters\OEM"
    if (-not (Test-Path $WlanParamPath)) { New-Item -Path $WlanParamPath -Force | Out-Null }
    
    try {
        New-ItemProperty -Path $WlanParamPath -Name "AllowScan" -Value 1 -PropertyType DWord -Force -ErrorAction SilentlyContinue | Out-Null
        Write-Log "Registry-Wert 'AllowScan' erfolgreich gesetzt."
    } catch {
        Write-Log "Fehler beim Setzen von AllowScan: ${ _ }" "ERROR"
    }

    # WLAN-Adapter Reset
    $adapters = Get-NetAdapter | Where-Object { $_.InterfaceDescription -like "*Wi-Fi*" -or $_.MediaType -eq "Native 802.11" }
    foreach ($adapter in $adapters) {
        Write-Log "Re-Initialisiere Adapter: $($adapter.Name)"
        Disable-NetAdapter -Name $adapter.Name -Confirm:$false
        Start-Sleep -Seconds 1
        Enable-NetAdapter -Name $adapter.Name -Confirm:$false
    }
}

# ==============================================================================
# 6. FINISH & GUI INTERACTION
# ==============================================================================
Write-Log "Aktualisiere System-Richtlinien..."
gpupdate /force | Out-Null

Write-Host "`n--- VORGANG ABGESCHLOSSEN ---" -ForegroundColor Cyan

if ($action -eq "Enable") {
    Write-Host "`nHINWEIS ZUR GUI:" -ForegroundColor Yellow
    Write-Host "Obwohl die Ortung im Hintergrund aktiv ist (Scanner funktionieren)," -ForegroundColor White
    Write-Host "zeigt die Windows-GUI eventuell noch 'Aus' an." -ForegroundColor White
    Write-Log "Öffne Windows-Einstellungsseite für den Benutzer..."
    
    # Öffne die Privacy-Location Seite direkt
    Start-Process "ms-settings:privacy-location"
    
    Write-Host "`nBITTE PRÜFEN SIE IN DER GEÖFFNETEN GUI:" -ForegroundColor Green
    Write-Host "1. 'Ortung für dieses Gerät' -> Ändern -> EIN" -ForegroundColor White
    Write-Host "2. 'Apps den Zugriff auf Ihren Standort erlauben' -> EIN" -ForegroundColor White
} else {
    Write-Host "Ortungsdienste sind DEAKTIVIERT." -ForegroundColor White
}

Write-Log "Skript erfolgreich beendet."
pause