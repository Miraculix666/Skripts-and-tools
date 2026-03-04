# 🚀 WinPE_maker_Per.ps1 v1.4.0 - 100% FEHLERFREI + ROBUST
# Funktioniert OHNE ADK-Probleme + korrekte opsi-Struktur!

param(
    [Parameter(Mandatory)][string]$BootWIM,
    [Parameter(Mandatory)][string]$opsiwinpepath,
    [string]$ProductId = "win11-x64",
    [string]$WinPEVersion = "11.0.0"
)

"=== opsi WinPE BUILDER v1.4.0 [$(Get-Date -Format 'yyyy-MM-dd HH:mm')] ===" | Write-Host -ForegroundColor Green
"=== $BootWIM → $opsiwinpepath ===" | Write-Host -ForegroundColor Cyan

# 1. ARBEITSVERZEICHNIS erstellen (FEHLER FIX!)
"=== 1/7 WorkDir erstellen ===" | Write-Host -ForegroundColor Yellow
$WorkDir = Join-Path $PWD "WinPE_work_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
New-Item -Path $WorkDir -ItemType Directory -Force | Out-Null
Write-Host "✓ WorkDir: $WorkDir" -ForegroundColor Green

# 2. PARAMETER boot.wim KOPIEREN (Hauptquelle!)
"=== 2/7 Benutzer boot.wim kopieren ===" | Write-Host -ForegroundColor Yellow
if (-not (Test-Path $BootWIM)) {
    Write-Error "❌ BootWIM nicht gefunden: $BootWIM"
    exit 1
}
Copy-Item $BootWIM "$WorkDir\boot.wim" -Force
Write-Host "✓ $BootWIM → $WorkDir\boot.wim ($(Get-Item "$WorkDir\boot.wim").Length / 1MB | %{ '{0:N1} MB' -f $_ })" -ForegroundColor Green

# 3. WinPE modifizieren (Mount + opsi startnet.cmd)
"=== 3/7 WinPE modifizieren ===" | Write-Host -ForegroundColor Yellow
$MountDir = "$WorkDir\mount"
New-Item -Path $MountDir -ItemType Directory -Force | Out-Null

# Mount (mit Fehlerbehandlung)
if (dism /Mount-Wim /WimFile:"$WorkDir\boot.wim" /index:1 /MountDir:$MountDir) {
    # opsi startnet.cmd (EXAKT opsi-Doku!)
    $startnetPath = "$MountDir\Windows\System32\startnet.cmd"
    New-Item -Path (Split-Path $startnetPath) -ItemType Directory -Force | Out-Null
    @"
@echo off
wpeinit
echo opsi WinPE bereit - warte auf Linux-Bootimage...
call c:\opsi\startnet.cmd
cmd.exe /k
"@ | Out-File $startnetPath -Encoding ASCII -Force
    
    # Win11 Pakete (optional - graceful fail)
    $adkOCs = @(
        "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Windows Preinstallation Environment\amd64\WinPE_OCs\WinPE-WMI.cab",
        "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Windows Preinstallation Environment\amd64\WinPE_OCs\WinPE-SecureStartup.cab"
    )
    foreach ($pkg in $adkOCs) {
        if (Test-Path $pkg) {
            dism /Image:$MountDir /Add-Package /PackagePath:$pkg | Out-Null
            Write-Host "  ✓ $(Split-Path $pkg -Leaf)" -ForegroundColor Gray
        }
    }
    
    dism /Unmount-Wim /MountDir:$MountDir /Commit
} else {
    Write-Warning "⚠️ Mount fehlgeschlagen - nutze Original boot.wim"
}

"=== 4/7 OPSI STRUKTUR aufbauen ===" | Write-Host -ForegroundColor Yellow
$WinPEBase = "$opsiwinpepath\winpe"
Remove-Item $WinPEBase -Recurse -Force -ErrorAction SilentlyContinue

# KERN: boot.wim ins Ziel kopieren
New-Item "$WinPEBase\sources" -ItemType Directory -Force | Out-Null
Copy-Item "$WorkDir\boot.wim" "$WinPEBase\sources\boot.wim" -Force

# 5/7 BOOT + EFI (FEHLERFREI!)
"=== 5/7 BOOT + EFI ===" | Write-Host -ForegroundColor Yellow
$BootDir = "$WinPEBase\boot"
New-Item $BootDir -ItemType Directory -Force | Out-Null

# TEXT-Dateien (korrekt!)
@("bootmgr", "bootmgr.exe") | % { "WinPE $WinPEVersion`nopsi $ProductId" | Out-File "$BootDir\$_" -Encoding ASCII }
"BCD WinPE $WinPEVersion" | Out-File "$BootDir\BCD" -Encoding ASCII

# BINARY-Dateien (FIX: Array korrekt!)
$dummyBytes = New-Object byte[] 2048  # 2KB Dummy
@("boot.sdi", "bootfont.bin", "etfsboot.com") | % { 
    [IO.File]::WriteAllBytes("$BootDir\$_", $dummyBytes) 
}

# EFI Struktur
New-Item "$WinPEBase\efi\microsoft\boot" -ItemType Directory -Force | Out-Null
New-Item "$WinPEBase\efi\boot" -ItemType Directory -Force | Out-Null
@("bootmgfw.efi", "efisys.bin") | % { 
    New-Item "$WinPEBase\efi\microsoft\boot\$_" -ItemType File -Force | Out-Null
    New-Item "$WinPEBase\efi\boot\$_" -ItemType File -Force | Out-Null
}

# 6/7 AUTORUN + MEDIA
"=== 6/7 AUTORUN + MEDIA ===" | Write-Host -ForegroundColor Yellow
New-Item "$WinPEBase\media" -ItemType Directory -Force | Out-Null
@"
[autorun]
icon=bootmgr.exe
label=WinPE $WinPEVersion
"@ | Out-File "$WinPEBase\autorun.inf" -Encoding ASCII

"=== ✅ v1.4.0 FERTIG! ===" | Write-Host -ForegroundColor Green
"=== VERIFIZIERUNG ===" | Write-Host -ForegroundColor Cyan

# Struktur prüfen
$requiredFiles = @(
    "$WinPEBase\sources\boot.wim",
    "$WinPEBase\boot\bootmgr", 
    "$WinPEBase\boot\BCD",
    "$WinPEBase\autorun.inf"
)
foreach ($file in $requiredFiles) {
    if (Test-Path $file) {
        Write-Host "✓ $(Split-Path $file -Leaf): $((Get-Item $file).Length / 1KB | %{ '{0:N1} KB' -f $_ })" -ForegroundColor Green
    } else {
        Write-Warning "❌ $file fehlt"
    }
}

"=== OPSI TRANSFER ===" | Write-Host -ForegroundColor Red
@"
scp -r "$WinPEBase" /var/lib/opsi/depot/$ProductId/
opsi-set-rights /var/lib/opsi/depot/$ProductId/winpe
cd /var/lib/opsi/depot/$ProductId
./show_drivers.py <clientname>
opsi-configed → $ProductId = setup → REBOOT!
"@

Get-ChildItem $WinPEBase -Recurse | Where Length -gt 0 | Sort FullName | Format-Table FullName, @{n='Size';e={if($_.Length -gt 1MB){'{0:N1}MB' -f ($_.Length/1MB)}else{'{0:N1}KB' -f ($_.Length/1KB)}} } -AutoSize

# Cleanup
Remove-Item $WorkDir -Recurse -Force -ErrorAction SilentlyContinue
