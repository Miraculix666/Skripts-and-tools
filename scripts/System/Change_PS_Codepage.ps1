# --- Automatische Adminrechte-Anforderung ---
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "WARNUNG: Skript läuft NICHT mit Adminrechten!" -ForegroundColor Yellow
    Write-Host "Starte Skript mit Adminrechten neu..." -ForegroundColor Yellow
    $arguments = "-NoExit -NoProfile -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`""
    Start-Process powershell.exe -ArgumentList $arguments -Verb RunAs
    Write-Host "Bitte warten Sie, während das Skript mit erhöhten Rechten startet..." -ForegroundColor Cyan
    exit
}

# --- Initialisierung und Systeminfo ---
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "PowerShell Codepage & Encoding Setup" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

Write-Host "[INFO] Systeminformationen:" -ForegroundColor Cyan
Write-Host "  - Windows Version: $([System.Environment]::OSVersion.VersionString)" -ForegroundColor White
Write-Host "  - PowerShell Version: $($PSVersionTable.PSVersion)" -ForegroundColor White
Write-Host "  - Benutzerkonto: $($env:USERNAME)" -ForegroundColor White
Write-Host "  - Admin-Status: JA (bestätigt)" -ForegroundColor Green
Write-Host "  - Aktuelles Encoding: $([Console]::OutputEncoding.WebName)" -ForegroundColor White
Write-Host ""

# --- 1. Codepage auf UTF-8 setzen ---
Write-Host "[SCHRITT 1/7] Setze Codepage auf UTF-8 (65001)..." -ForegroundColor Cyan
try {
    $currentCodepage = chcp
    Write-Host "  - Aktuelle Codepage vor Änderung: $currentCodepage" -ForegroundColor Gray
    
    chcp 65001 2>&1 | Out-Null
    Start-Sleep -Milliseconds 500
    
    $newCodepage = chcp
    Write-Host "  - Neue Codepage: $newCodepage" -ForegroundColor Gray
    
    if ($newCodepage -match "65001") {
        Write-Host "  ✓ Codepage erfolgreich auf UTF-8 (65001) gesetzt." -ForegroundColor Green
    } else {
        Write-Host "  ✗ Codepage-Änderung möglicherweise nicht erfolgreich." -ForegroundColor Yellow
        Write-Host "    (In Windows PE kann dies erwartet sein)" -ForegroundColor Gray
    }
} catch {
    Write-Host "  ✗ Fehler beim Setzen der Codepage: $_" -ForegroundColor Red
}
Write-Host ""

# --- 2. Umgebungsvariable setzen ---
Write-Host "[SCHRITT 2/7] Setze Umgebungsvariablen..." -ForegroundColor Cyan
try {
    $env:LC_ALL = "de_DE.UTF-8"
    Write-Host "  - Setze LC_ALL = '$env:LC_ALL'" -ForegroundColor Gray
    
    if ($env:LC_ALL -eq "de_DE.UTF-8") {
        Write-Host "  ✓ Umgebungsvariable LC_ALL erfolgreich gesetzt." -ForegroundColor Green
    } else {
        Write-Host "  ✗ Fehler beim Setzen von LC_ALL" -ForegroundColor Red
    }
    
    # Zusätzliche Encoding-Umgebungsvariablen
    $env:PYTHONIOENCODING = "utf-8"
    Write-Host "  - Setze PYTHONIOENCODING = 'utf-8'" -ForegroundColor Gray
    Write-Host "  ✓ Weitere Umgebungsvariablen konfiguriert." -ForegroundColor Green
} catch {
    Write-Host "  ✗ Fehler beim Setzen der Umgebungsvariablen: $_" -ForegroundColor Red
}
Write-Host ""

# --- 3. Konsolen-Encoding setzen ---
Write-Host "[SCHRITT 3/7] Konfiguriere Konsolen-Encoding..." -ForegroundColor Cyan
try {
    Write-Host "  - Altes OutputEncoding: $([Console]::OutputEncoding.WebName)" -ForegroundColor Gray
    Write-Host "  - Altes InputEncoding: $([Console]::InputEncoding.WebName)" -ForegroundColor Gray
    
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    [Console]::InputEncoding = [System.Text.Encoding]::UTF8
    
    Write-Host "  - Neues OutputEncoding: $([Console]::OutputEncoding.WebName)" -ForegroundColor Gray
    Write-Host "  - Neues InputEncoding: $([Console]::InputEncoding.WebName)" -ForegroundColor Gray
    
    if ([Console]::OutputEncoding.WebName -eq "utf-8" -and [Console]::InputEncoding.WebName -eq "utf-8") {
        Write-Host "  ✓ Konsolen-Encoding erfolgreich auf UTF-8 gesetzt." -ForegroundColor Green
    } else {
        Write-Host "  ✗ Konsolen-Encoding konnte nicht vollständig gesetzt werden." -ForegroundColor Red
    }
} catch {
    Write-Host "  ✗ Fehler beim Setzen des Konsolen-Encodings: $_" -ForegroundColor Red
}
Write-Host ""

# --- 4. Sprache, Layout und Zeitzone setzen ---
Write-Host "[SCHRITT 4/7] Konfiguriere Sprache, Layout und Zeitzone..." -ForegroundColor Cyan
try {
    # Versuche WinUI Language Override (kann in PE fehlschlagen)
    try {
        Write-Host "  - Setze WinUILanguageOverride auf 'de-DE'..." -ForegroundColor Gray
        Set-WinUILanguageOverride -Language "de-DE" -ErrorAction Stop
        Write-Host "    ✓ WinUILanguageOverride gesetzt" -ForegroundColor Green
    } catch {
        Write-Host "    ⚠ WinUILanguageOverride nicht verfügbar (erwartet in PE): $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    # Versuche User Language List zu setzen
    try {
        Write-Host "  - Setze WinUserLanguageList auf 'de-DE'..." -ForegroundColor Gray
        Set-WinUserLanguageList -LanguageList "de-DE" -Force -ErrorAction Stop
        Write-Host "    ✓ WinUserLanguageList gesetzt" -ForegroundColor Green
    } catch {
        Write-Host "    ⚠ WinUserLanguageList nicht verfügbar: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    # Setze Culture
    try {
        Write-Host "  - Setze Culture auf 'de-DE'..." -ForegroundColor Gray
        Set-Culture -CultureInfo "de-DE" -ErrorAction Stop
        Write-Host "    ✓ Culture gesetzt zu: $(Get-Culture)" -ForegroundColor Green
    } catch {
        Write-Host "    ✗ Fehler beim Setzen der Culture: $_" -ForegroundColor Red
    }
    
    # Setze System Locale
    try {
        Write-Host "  - Setze WinSystemLocale auf 'de-DE'..." -ForegroundColor Gray
        Set-WinSystemLocale -SystemLocale "de-DE" -ErrorAction Stop
        Write-Host "    ✓ WinSystemLocale gesetzt" -ForegroundColor Green
    } catch {
        Write-Host "    ⚠ WinSystemLocale konnte nicht gesetzt werden: $_" -ForegroundColor Yellow
    }
    
    # Setze Zeitzone
    try {
        Write-Host "  - Setze Zeitzone auf 'W. Europe Standard Time'..." -ForegroundColor Gray
        Set-TimeZone -Id "W. Europe Standard Time" -ErrorAction Stop
        Write-Host "    ✓ Zeitzone gesetzt zu: $(Get-TimeZone)" -ForegroundColor Green
    } catch {
        Write-Host "    ⚠ Zeitzone konnte nicht gesetzt werden: $_" -ForegroundColor Yellow
    }
    
    Write-Host "  ✓ Sprache- und Lokalisierungseinstellungen konfiguriert." -ForegroundColor Green
} catch {
    Write-Host "  ✗ Fehler bei Sprach-/Lokalisierungseinstellungen: $_" -ForegroundColor Red
}
Write-Host ""

# --- 5. Systemweite OEM-Codepage setzen (erfordert Adminrechte) ---
Write-Host "[SCHRITT 5/7] Setze systemweite OEM-Codepage in Registry..." -ForegroundColor Cyan
try {
    $regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Nls\CodePage"
    Write-Host "  - Registry-Pfad: $regPath" -ForegroundColor Gray
    
    # Überprüfe aktuelle Werte
    Write-Host "  - Aktuelle Registry-Werte:" -ForegroundColor Gray
    try {
        $oemcp = Get-ItemProperty -Path $regPath -Name "OEMCP" -ErrorAction SilentlyContinue
        if ($oemcp) {
            Write-Host "    • OEMCP: $($oemcp.OEMCP)" -ForegroundColor Gray
        }
        $acp = Get-ItemProperty -Path $regPath -Name "ACP" -ErrorAction SilentlyContinue
        if ($acp) {
            Write-Host "    • ACP: $($acp.ACP)" -ForegroundColor Gray
        }
        $maccp = Get-ItemProperty -Path $regPath -Name "MACCP" -ErrorAction SilentlyContinue
        if ($maccp) {
            Write-Host "    • MACCP: $($maccp.MACCP)" -ForegroundColor Gray
        }
    } catch {
        Write-Host "    ⚠ Konnte aktuelle Werte nicht auslesen: $_" -ForegroundColor Yellow
    }
    
    # Setze neue Werte
    Write-Host "  - Setze neue Werte..." -ForegroundColor Gray
    New-ItemProperty -Path $regPath -Name "OEMCP" -Value "65001" -PropertyType String -Force -ErrorAction Stop | Out-Null
    Write-Host "    ✓ OEMCP = 65001" -ForegroundColor Green
    
    New-ItemProperty -Path $regPath -Name "ACP" -Value "65001" -PropertyType String -Force -ErrorAction Stop | Out-Null
    Write-Host "    ✓ ACP = 65001" -ForegroundColor Green
    
    New-ItemProperty -Path $regPath -Name "MACCP" -Value "65001" -PropertyType String -Force -ErrorAction Stop | Out-Null
    Write-Host "    ✓ MACCP = 65001" -ForegroundColor Green
    
    Write-Host "  ✓ Systemweite OEM-Codepage auf UTF-8 gesetzt (Neustart erforderlich)." -ForegroundColor Green
} catch {
    Write-Host "  ✗ Fehler beim Setzen der systemweiten OEM-Codepage: $_" -ForegroundColor Red
}
Write-Host ""

# --- 6. PowerShell-Profil dauerhaft anpassen ---
Write-Host "[SCHRITT 6/7] Passe PowerShell-Profil an..." -ForegroundColor Cyan
try {
    $profilePath = $PROFILE.CurrentUserCurrentHost
    Write-Host "  - Profil-Pfad: $profilePath" -ForegroundColor Gray
    
    if (-not (Test-Path -Path $profilePath)) {
        Write-Host "  - Profil existiert nicht, erstelle neues..." -ForegroundColor Gray
        $profileDir = Split-Path -Path $profilePath -Parent
        if (-not (Test-Path -Path $profileDir)) {
            New-Item -ItemType Directory -Path $profileDir -Force | Out-Null
            Write-Host "    ✓ Profil-Verzeichnis erstellt: $profileDir" -ForegroundColor Green
        }
        New-Item -ItemType File -Path $profilePath -Force | Out-Null
        Write-Host "    ✓ Profil-Datei erstellt" -ForegroundColor Green
    } else {
        Write-Host "  - Profil existiert bereits" -ForegroundColor Gray
    }
    
    $encodingLines = @"
# UTF-8 Encoding für PowerShell (automatisch hinzugefügt)
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
"@
    
    $currentProfileContent = Get-Content -Path $profilePath -Raw -ErrorAction SilentlyContinue
    if ([string]::IsNullOrEmpty($currentProfileContent)) {
        Write-Host "  - Profil ist leer, schreibe Encoding-Einstellungen..." -ForegroundColor Gray
        Set-Content -Path $profilePath -Value $encodingLines -Force
        Write-Host "    ✓ Encoding-Einstellungen geschrieben" -ForegroundColor Green
    } elseif ($currentProfileContent -notmatch '\[Console\]::OutputEncoding') {
        Write-Host "  - Encoding-Einstellungen nicht im Profil, füge hinzu..." -ForegroundColor Gray
        Add-Content -Path $profilePath -Value "`n$encodingLines"
        Write-Host "    ✓ Encoding-Einstellungen hinzugefügt" -ForegroundColor Green
    } else {
        Write-Host "  - Encoding-Einstellungen sind bereits im Profil vorhanden" -ForegroundColor Yellow
    }
    
    Write-Host "  ✓ PowerShell-Profil erfolgreich angepasst." -ForegroundColor Green
} catch {
    Write-Host "  ✗ Fehler beim Anpassen des PowerShell-Profils: $_" -ForegroundColor Red
}
Write-Host ""

# --- 7. Zusammenfassung und Registry-Überprüfung ---
Write-Host "[SCHRITT 7/7] Zusammenfassung und Überprüfung..." -ForegroundColor Cyan
Write-Host "  - Aktuelle Systemeinstellungen:" -ForegroundColor Gray
Write-Host "    • Consolencodierung: $([Console]::OutputEncoding.WebName)" -ForegroundColor White
Write-Host "    • Culture: $(Get-Culture)" -ForegroundColor White
Write-Host "    • Zeitzone: $(Get-TimeZone)" -ForegroundColor White
Write-Host "    • LC_ALL: $($env:LC_ALL)" -ForegroundColor White

try {
    $regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Nls\CodePage"
    $oemcp = Get-ItemProperty -Path $regPath -Name "OEMCP" -ErrorAction SilentlyContinue
    Write-Host "    • Registry OEMCP: $($oemcp.OEMCP)" -ForegroundColor White
} catch {
    Write-Host "    • Registry OEMCP: (nicht lesbar)" -ForegroundColor Gray
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Konfiguration abgeschlossen!" -ForegroundColor Green
Write-Host "========================================`n" -ForegroundColor Cyan

Write-Host "WICHTIG: Neustart erforderlich für:" -ForegroundColor Yellow
Write-Host "  - Systemweite OEM-Codepage-Änderungen" -ForegroundColor White
Write-Host "  - Windows PE Locale-Änderungen (falls zutreffend)" -ForegroundColor White
Write-Host ""

Write-Host "Die Einstellungen sind jetzt konfiguriert." -ForegroundColor Green
Write-Host "Bitte führen Sie bei Bedarf einen Neustart durch." -ForegroundColor Cyan
Write-Host ""

# --- Konsole offen lassen - INTERACTIVE MODE ---
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "INTERAKTIVE KONSOLE - TERMINAL BLEIBT OFFEN" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Sie können jetzt weitere Befehle eingeben oder 'exit' eintippen zum Beenden." -ForegroundColor Cyan
Write-Host "Geben Sie 'Test-Codepage' ein, um die aktuellen Einstellungen zu überprüfen." -ForegroundColor Cyan
Write-Host ""

# Definiere Hilfsfunktion für Tests
function Test-Codepage {
    Write-Host "`nCodepage & Encoding Test:" -ForegroundColor Cyan
    Write-Host "  Aktuelle Codepage: $(chcp)" -ForegroundColor White
    Write-Host "  OutputEncoding: $([Console]::OutputEncoding.WebName)" -ForegroundColor White
    Write-Host "  InputEncoding: $([Console]::InputEncoding.WebName)" -ForegroundColor White
    Write-Host "  Culture: $(Get-Culture)" -ForegroundColor White
    Write-Host "  Timezone: $(Get-TimeZone)" -ForegroundColor White
    Write-Host "";
}

# Keep session open - Prevent window from closing
while ($true) {
    Write-Host "PS> " -NoNewline -ForegroundColor Green
    $input = Read-Host;
    
    if ($input -eq "exit" -or $input -eq "quit") {
        Write-Host "Auf Wiedersehen!" -ForegroundColor Cyan
        break;
    } elseif ($input -eq "Test-Codepage" -or $input -eq "test") {
        Test-Codepage;
    } elseif ([string]::IsNullOrWhiteSpace($input)) {
        continue;
    } else {
        try {
            Invoke-Expression $input;
        } catch {
            Write-Host "Fehler bei der Ausführung: $_" -ForegroundColor Red;
        }
    }
}

Write-Host "";
Write-Host "Konsole wird geschlossen..." -ForegroundColor Yellow;
Start-Sleep -Seconds 2;