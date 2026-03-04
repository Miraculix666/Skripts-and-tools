# BCD_QuickFix.ps1
# Repariert BCD in bestehendem WinPE ohne kompletten Rebuild
# Für Error 0xc0000098 (BCD Missing/Corrupt)

Param(
    [Parameter(Mandatory=$true)]
    [string]$WinPEPath,  # Pfad zum winpe Ordner z.B. "Y:\win11-x64\winpe"
    
    [Parameter(Mandatory=$false)]
    [string]$ISOPath     # Optional: ISO als Quelle für BCDTemplate
)

$ErrorActionPreference = "Stop"

Write-Host "`n===============================================" -ForegroundColor Cyan
Write-Host "   BCD Quick-Fix für Error 0xc0000098        " -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan

if (-not (Test-Path $WinPEPath)) {
    throw "WinPE-Pfad nicht gefunden: $WinPEPath"
}

Write-Host "[CHECK] Prüfe Struktur..." -ForegroundColor Yellow

# Prüfe ob boot.wim existiert
$bootWim = Join-Path $WinPEPath "sources\boot.wim"
if (-not (Test-Path $bootWim)) {
    throw "boot.wim nicht gefunden: $bootWim"
}
Write-Host " [OK] boot.wim gefunden" -ForegroundColor Green

# Prüfe ob boot.sdi existiert
$bootSdi = Join-Path $WinPEPath "Boot\boot.sdi"
if (-not (Test-Path $bootSdi)) {
    Write-Warning "boot.sdi fehlt! Versuche aus ISO zu kopieren..."
    if ($ISOPath -and (Test-Path $ISOPath)) {
        $iso = Mount-DiskImage $ISOPath -PassThru
        $drive = ($iso | Get-Volume).DriveLetter
        Copy-Item "${drive}:\Boot\boot.sdi" $bootSdi -Force
        Dismount-DiskImage $ISOPath | Out-Null
        Write-Host " [OK] boot.sdi von ISO kopiert" -ForegroundColor Green
    } else {
        throw "boot.sdi fehlt und keine ISO angegeben!"
    }
}

Write-Host "`n[BCD] Repariere Boot Configuration..." -ForegroundColor Yellow

$ramdiskGuid = "{7619dcc8-fafe-11d9-b411-000476eba25f}"

# === BIOS BCD ===
$biosBCD = Join-Path $WinPEPath "Boot\BCD"
$biosBCDTemplate = Join-Path $WinPEPath "Boot\BCDTemplate"

Write-Host " [BIOS] Erstelle BCD..." -ForegroundColor Cyan

# Backup alte BCD
if (Test-Path $biosBCD) {
    Copy-Item $biosBCD "$biosBCD.backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')" -Force
    Remove-Item $biosBCD -Force
}

# Verwende Template falls vorhanden
if (Test-Path $biosBCDTemplate) {
    Write-Host "  [TEMPLATE] Verwende BCDTemplate..." -ForegroundColor Gray
    Copy-Item $biosBCDTemplate $biosBCD -Force
} else {
    Write-Host "  [CREATE] Erstelle neue BCD..." -ForegroundColor Gray
    
    # Erstelle leere BCD-Datei
    $result = & bcdedit.exe /createstore $biosBCD 2>&1
    Write-Verbose "bcdedit createstore: $result"
    
    # Prüfe ob BCD erstellt wurde
    if (-not (Test-Path $biosBCD)) {
        Write-Warning "bcdedit konnte BCD nicht erstellen! Verwende Alternative..."
        
        # Alternative: Kopiere von ISO falls vorhanden
        if ($ISOPath -and (Test-Path $ISOPath)) {
            $iso = Mount-DiskImage $ISOPath -PassThru
            $drive = ($iso | Get-Volume).DriveLetter
            $isoBCD = "${drive}:\Boot\BCD"
            if (Test-Path $isoBCD) {
                Copy-Item $isoBCD $biosBCD -Force
                Write-Host "  [OK] BCD von ISO kopiert" -ForegroundColor Green
            }
            Dismount-DiskImage $ISOPath | Out-Null
        }
        
        # Wenn immer noch nicht vorhanden, Fehler
        if (-not (Test-Path $biosBCD)) {
            throw "BCD konnte nicht erstellt werden! Führen Sie das Skript als Administrator aus."
        }
    }
    
    # Boot Manager
    $bootmgrGuid = "{9dea862c-5cdd-4e70-acc1-f32b344d4795}"
    & bcdedit.exe /store $biosBCD /create $bootmgrGuid /d "Windows Boot Manager" 2>&1 | Out-Null
    & bcdedit.exe /store $biosBCD /set $bootmgrGuid device boot 2>&1 | Out-Null
    & bcdedit.exe /store $biosBCD /set $bootmgrGuid path "\Boot\bootmgfw.efi" 2>&1 | Out-Null
    & bcdedit.exe /store $biosBCD /set $bootmgrGuid displayorder "{default}" 2>&1 | Out-Null
    & bcdedit.exe /store $biosBCD /set $bootmgrGuid timeout 1 2>&1 | Out-Null
    
    # OS Loader
    & bcdedit.exe /store $biosBCD /create "{default}" /d "Microsoft Windows PE" /application osloader 2>&1 | Out-Null
}

# Passe BCD an
& bcdedit.exe /store $biosBCD /set "{default}" device "ramdisk=[boot]\sources\boot.wim,$ramdiskGuid" 2>&1 | Out-Null
& bcdedit.exe /store $biosBCD /set "{default}" osdevice "ramdisk=[boot]\sources\boot.wim,$ramdiskGuid" 2>&1 | Out-Null
& bcdedit.exe /store $biosBCD /set "{default}" path "\windows\system32\boot\winload.exe" 2>&1 | Out-Null
& bcdedit.exe /store $biosBCD /set "{default}" systemroot "\windows" 2>&1 | Out-Null
& bcdedit.exe /store $biosBCD /set "{default}" detecthal Yes 2>&1 | Out-Null
& bcdedit.exe /store $biosBCD /set "{default}" winpe Yes 2>&1 | Out-Null
& bcdedit.exe /store $biosBCD /set "{default}" nointegritychecks Yes 2>&1 | Out-Null
& bcdedit.exe /store $biosBCD /set "{default}" testsigning Yes 2>&1 | Out-Null
& bcdedit.exe /store $biosBCD /set "{default}" bootmenupolicy Legacy 2>&1 | Out-Null

# Ramdisk Device
$ramdiskExists = & bcdedit.exe /store $biosBCD /enum all 2>&1 | Select-String $ramdiskGuid
if (-not $ramdiskExists) {
    & bcdedit.exe /store $biosBCD /create $ramdiskGuid /d "Ramdisk Options" /device 2>&1 | Out-Null
}
& bcdedit.exe /store $biosBCD /set $ramdiskGuid ramdisksdidevice boot 2>&1 | Out-Null
& bcdedit.exe /store $biosBCD /set $ramdiskGuid ramdisksdipath "\Boot\boot.sdi" 2>&1 | Out-Null

# Entferne Read-Only
if (Test-Path $biosBCD) {
    Set-ItemProperty $biosBCD -Name IsReadOnly -Value $false
    Write-Host "  [OK] BIOS BCD erstellt" -ForegroundColor Green
} else {
    throw "FEHLER: BIOS BCD wurde nicht erstellt!"
}

# === UEFI BCD ===
$uefiBCD = Join-Path $WinPEPath "EFI\Microsoft\Boot\BCD"

Write-Host " [UEFI] Erstelle BCD..." -ForegroundColor Cyan

# Backup
if (Test-Path $uefiBCD) {
    Copy-Item $uefiBCD "$uefiBCD.backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')" -Force
    Remove-Item $uefiBCD -Force
}

# Kopiere BIOS BCD als Basis
Copy-Item $biosBCD $uefiBCD -Force

# Passe für UEFI an
& bcdedit.exe /store $uefiBCD /set "{default}" path "\windows\system32\boot\winload.efi" 2>&1 | Out-Null
& bcdedit.exe /store $uefiBCD /set "{9dea862c-5cdd-4e70-acc1-f32b344d4795}" path "\EFI\Microsoft\Boot\bootmgfw.efi" 2>&1 | Out-Null

if (Test-Path $uefiBCD) {
    Set-ItemProperty $uefiBCD -Name IsReadOnly -Value $false
    Write-Host "  [OK] UEFI BCD erstellt" -ForegroundColor Green
} else {
    throw "FEHLER: UEFI BCD wurde nicht erstellt!"
}

# === VALIDIERUNG ===
Write-Host "`n[VERIFY] Validiere BCD..." -ForegroundColor Yellow

$biosCheck = & bcdedit.exe /store $biosBCD /enum all 2>&1
if ($biosCheck -match "ramdisk=\[boot\]\\sources\\boot.wim") {
    Write-Host " [OK] BIOS BCD korrekt konfiguriert" -ForegroundColor Green
} else {
    Write-Warning " [WARN] BIOS BCD möglicherweise fehlerhaft!"
}

$uefiCheck = & bcdedit.exe /store $uefiBCD /enum all 2>&1
if ($uefiCheck -match "ramdisk=\[boot\]\\sources\\boot.wim") {
    Write-Host " [OK] UEFI BCD korrekt konfiguriert" -ForegroundColor Green
} else {
    Write-Warning " [WARN] UEFI BCD möglicherweise fehlerhaft!"
}

Write-Host "`n===============================================" -ForegroundColor Cyan
Write-Host "   BCD-Reparatur abgeschlossen                " -ForegroundColor Green
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host "`nNächste Schritte:" -ForegroundColor Yellow
Write-Host " 1. Falls OPSI-Depot: opsi-set-rights ausführen" -ForegroundColor Gray
Write-Host " 2. Secure Boot im BIOS deaktivieren" -ForegroundColor Gray
Write-Host " 3. Test-Boot durchführen" -ForegroundColor Gray
Write-Host "`nBCD-Dumps für Debugging:" -ForegroundColor Cyan
Write-Host " BIOS: bcdedit /store '$biosBCD' /enum all" -ForegroundColor Gray
Write-Host " UEFI: bcdedit /store '$uefiBCD' /enum all" -ForegroundColor Gray
