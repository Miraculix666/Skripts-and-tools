# --- Automatische Adminrechte-Anforderung ---
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # Skript neu mit Adminrechten und -NoExit starten, damit das Fenster offen bleibt
    $arguments = "-NoExit -File `"$($MyInvocation.MyCommand.Path)`""
    Start-Process powershell.exe -ArgumentList $arguments -Verb RunAs
    exit
}

# --- 1. Codepage auf UTF-8 setzen ---
chcp 65001 | Out-Null
if ((chcp) -match "65001") {
    Write-Host "Codepage erfolgreich auf UTF-8 gesetzt." -ForegroundColor Green
} else {
    Write-Host "Fehler beim Setzen der Codepage!" -ForegroundColor Red
}

# --- 2. Umgebungsvariable setzen ---
$env:LC_ALL = "de_DE.UTF-8"
if ($env:LC_ALL -eq "de_DE.UTF-8") {
    Write-Host "Umgebungsvariable LC_ALL gesetzt." -ForegroundColor Green
} else {
    Write-Host "Fehler beim Setzen der Umgebungsvariable LC_ALL!" -ForegroundColor Red
}

# --- 3. Konsolen-Encoding setzen ---
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8
if ([Console]::OutputEncoding.WebName -eq "utf-8" -and [Console]::InputEncoding.WebName -eq "utf-8") {
    Write-Host "Konsolen-Encoding erfolgreich auf UTF-8 gesetzt." -ForegroundColor Green
} else {
    Write-Host "Fehler beim Setzen des Konsolen-Encodings!" -ForegroundColor Red
}

# --- 4. Sprache, Layout und Zeitzone setzen ---
try {
    Set-WinUILanguageOverride -Language "de-DE"
    Set-WinUserLanguageList -LanguageList "de-DE" -Force
    Set-Culture -CultureInfo "de-DE"
    Set-WinSystemLocale -SystemLocale "de-DE"
    Set-TimeZone -Id "W. Europe Standard Time"
    Write-Host "Sprache, Layout und Zeitzone erfolgreich auf Deutsch/Deutschland gesetzt." -ForegroundColor Green
} catch {
    Write-Host "Fehler beim Setzen der Sprache, des Layouts oder der Zeitzone!" -ForegroundColor Red
}

# --- 5. Systemweite OEM-Codepage setzen (erfordert Adminrechte) ---
try {
    New-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Nls\CodePage" -Name "OEMCP" -Value "65001" -Force -ErrorAction Stop
    Write-Host "Systemweite OEM-Codepage erfolgreich auf UTF-8 gesetzt (Neustart erforderlich)." -ForegroundColor Green
} catch {
    Write-Host "Fehler beim Setzen der systemweiten OEM-Codepage (Adminrechte nötig)!" -ForegroundColor Yellow
}

# --- 6. PowerShell-Profil dauerhaft anpassen ---
$profilePath = $PROFILE.CurrentUserCurrentHost
if (-not (Test-Path -Path $profilePath)) {
    New-Item -ItemType File -Path $profilePath -Force | Out-Null
}
$encodingLines = @"
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
"@
$currentProfileContent = Get-Content -Path $profilePath -Raw
if ($currentProfileContent -notmatch '\[Console\]::OutputEncoding') {
    Add-Content -Path $profilePath -Value $encodingLines
    Write-Host "Kodierungseinstellungen dauerhaft im PowerShell-Profil gespeichert." -ForegroundColor Green
} else {
    Write-Host "Kodierungseinstellungen sind bereits im PowerShell-Profil vorhanden." -ForegroundColor Yellow
}

Write-Host "`nAlle Änderungen abgeschlossen. Bitte ggf. das System neu starten!" -ForegroundColor Cyan

# --- Konsole offen lassen ---
Write-Host ""
Read-Host "Drücken Sie die Eingabetaste, um das Fenster zu schließen ..."
