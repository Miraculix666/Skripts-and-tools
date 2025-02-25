#Somtimes Powershell has a wrong Codepage Setting, this script can correct this

# Setzt die Codepage auf UTF-8
chcp 65001

# Setzt die Umgebungsvariable auf UTF-8
$env:LC_ALL = "de_DE.UTF-8"

# Setzt die Konsolenausgabe auf UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Set locale to German (Germany)
Set-WinUILanguageOverride -Language "de-DE"
Set-WinUserLanguageList -LanguageList "de-DE" -Force
Set-Culture -CultureInfo "de-DE"
Set-WinSystemLocale -SystemLocale "de-DE"

# Set console output encoding to UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8


###General German localisation

# Set system locale to German
Set-WinSystemLocale -SystemLocale "de-DE"

# Set input method and keyboard layout to German
Set-WinUILanguageOverride -Language "de-DE"
Set-WinUserLanguageList -LanguageList "de-DE" -Force

# Set region format to Germany
Set-Culture -CultureInfo "de-DE"

# Optionally, set time zone to Central European Time
Set-TimeZone -Id "W. Europe Standard Time"


# Optional: Change system default encoding (requires admin privileges and a restart)
New-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Nls\CodePage" -Name "OEMCP" -Value "65001" -Force

Write-Host "Sprache und Layout wurden erfolgreich auf Deutsch umgestellt." -ForegroundColor Green

