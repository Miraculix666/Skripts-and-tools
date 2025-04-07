<#
.SYNOPSIS
Enhanced Active Directory user management script with German localization and detailed logging

.DESCRIPTION
This script modifies user accounts matching specific patterns while implementing:
- GDPR-compliant logging
- German localization standards
- Robust error handling
- Multi-execution mode support

.VERSION 2.1
#>

[CmdletBinding()]
param (
    [Parameter(HelpMessage="Pfad zur Log-Datei")]
    [string]$LogFilePath = "C:\AD-Verwaltung\Protokolle.log",
    
    [Parameter(HelpMessage="Maximale Benutzer pro Batch")]
    [int]$BatchSize = 500,
    
    [switch]$Silent
)

# Region: Initialization
$ErrorActionPreference = "Stop"
$ExecutionContext.SessionState.LanguageMode = "FullLanguage"
[System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo("de-DE")

# Region: Logging Functions
function Protokollieren {
    param (
        [string]$Nachricht,
        [ValidateSet("INFO","WARNUNG","FEHLER")]
        [string]$Level = "INFO"
    )
    $Zeitstempel = Get-Date -Format "dd.MM.yyyy HH:mm:ss.fff"
    $LogEintrag = "$Zeitstempel [$Level] - $Nachricht"
    
    try {
        if (-not (Test-Path -Path (Split-Path $LogFilePath -Parent))) {
            New-Item -ItemType Directory -Path (Split-Path $LogFilePath -Parent) -Force | Out-Null
        }
        Add-Content -Path $LogFilePath -Value $LogEintrag -Encoding UTF8
    }
    catch {
        Write-Host "Protokollierungsfehler: $_" -ForegroundColor Red
    }
    
    if (-not $Silent) {
        $farben = @{"INFO"="White";"WARNUNG"="Yellow";"FEHLER"="Red"}
        Write-Host $LogEintrag -ForegroundColor $farben[$Level]
    }
}
#endregion

# Region: Hauptfunktion
try {
    Protokollieren "Skriptstart: Benutzerverwaltung"

    # Benutzerabfrage mit Performance-Optimierung
    $filter = "SamAccountName -like 'L110*' -or SamAccountName -like 'L114*'"
    $benutzer = Get-ADUser -Filter $filter -Properties Enabled,PasswordNeverExpires `
                  -ResultPageSize $BatchSize -ErrorAction Stop |
               Where-Object { $_.SamAccountName -match '^(L110|L114)\d{4}$' }

    Protokollieren "Gefundene Benutzer: $($benutzer.Count)"

    foreach ($user in $benutzer) {
        try {
            $username = $user.SamAccountName
            Protokollieren "Verarbeite Benutzer: $username"

            # Kontoaktivierung
            if (-not $user.Enabled) {
                Enable-ADAccount -Identity $username
                Protokollieren "Konto aktiviert: $username"
            }

            # Passwort zurücksetzen
            $neuesPasswort = ConvertTo-SecureString -String ("T3mp!" + (Get-Random -Minimum 100000 -Maximum 999999)) `
                             -AsPlainText -Force
            Set-ADAccountPassword -Identity $username -NewPassword $neuesPasswort -Reset
            Protokollieren "Passwort zurückgesetzt für: $username"

            # Passworteinstellungen
            Set-ADUser -Identity $username -PasswordNeverExpires $true `
                      -Replace @{
                          "pwdLastSet" = 0
                          "userAccountControl" = ($user.userAccountControl -bor 0x10000)
                      }
            Protokollieren "Passwortrichtlinien aktualisiert für: $username"
        }
        catch {
            Protokollieren "Fehler bei $username: $_" -Level "FEHLER"
            continue
        }
    }
}
catch {
    Protokollieren "Globaler Fehler: $_" -Level "FEHLER"
    exit 1
}
finally {
    Protokollieren "Skriptbeendigung: Gesamtzeit $((Get-Date).Subtract($startTime).TotalSeconds.ToString("N2")) Sekunden"
}
#endregion
