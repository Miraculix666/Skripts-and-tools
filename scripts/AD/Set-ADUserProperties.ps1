<#
.SYNOPSIS
    Ändert den Status der Eigenschaft 'PasswordChangeRequired' für Active Directory-Benutzer.
    Standardmäßig wird 'PasswordChangeRequired' auf $false gesetzt, d.h. der Benutzer muss sein Passwort nicht beim nächsten Login ändern.

.DESCRIPTION
    Dieses Skript durchsucht Active Directory nach Benutzernamen, die einem bestimmten Muster entsprechen.
    Es kann optional auf eine bestimmte Organisationseinheit (OU) beschränkt werden.
    Vor jeder Änderung wird standardmäßig eine Bestätigung abgefragt.
    Die Ausführung erfolgt standardmäßig mit detaillierter Ausgabe (verbose).

.PARAMETER SearchPattern
    Das Suchmuster für die Benutzernamen. Platzhalter wie '*' können verwendet werden.
    Beispiel: 'Hans*' findet Benutzer wie 'Hans Müller', 'Hansi', etc.

.PARAMETER OrganizationalUnit
    Der Distinguished Name (DN) der Organisationseinheit, in der die Benutzer gesucht werden sollen.
    Wenn nicht angegeben, wird im gesamten Active Directory gesucht.
    Beispiel: 'OU=Benutzer,DC=contoso,DC=com'

.PARAMETER ConfirmAction
    Ein Schalter-Parameter. Wenn gesetzt, wird für jeden gefundenen Benutzer eine Bestätigungsabfrage angezeigt,
    bevor die Änderung durchgeführt wird. Standardmäßig aktiviert.
    Um diese Abfrage zu überspringen, setzen Sie '-ConfirmAction:$false'.

.PARAMETER Silent
    Ein Schalter-Parameter. Wenn gesetzt, wird die detaillierte Ausgabe (verbose) unterdrückt.
    Nur kritische Fehler und grundlegende Informationen werden angezeigt.

.NOTES
    Autor: PS-Coding
    Version: 1.1 (Korrigierter Syntaxfehler)
    Datum: 2025-05-27

    Anforderungen:
    - PowerShell Version 5.1 oder höher.
    - Active Directory PowerShell-Modul.
    - Ausreichende Berechtigungen im Active Directory, um Benutzerattribute zu ändern.

    AI-disocovered Sources:
    - Get-ADUser: https://learn.microsoft.com/en-us/powershell/module/activedirectory/get-aduser
    - Set-ADUser: https://learn.microsoft.com/en-us/powershell/module/activedirectory/set-aduser
    - Write-Verbose: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/write-verbose
    - Bestätigungslogik in PowerShell: https://learn.microsoft.com/en-us/powershell/scripting/developer/cmdlet/requesting-confirmation

    User-provided Sources:
    - Specific requirements regarding search pattern, OU scope, no password reset, interactive confirmation,
      success verification, German localization, and verbose output.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$SearchPattern,

    [Parameter(Mandatory = $false, Position = 1)]
    [string]$OrganizationalUnit,

    [Parameter(Mandatory = $false)]
    [bool]$ConfirmAction = $true, # Standardmäßig Bestätigung anfordern

    [Parameter(Mandatory = $false)]
    [switch]$Silent # Standardmäßig verbose, Silent unterdrückt es
)

# Region: Globale Variablen und Lokalisierungseinstellungen
# -----------------------------------------------------------
# Deutsche Lokalisierung für Datums- und Zeitformate
$CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo("de-DE")

# Für CSV-Export, falls benötigt, sonst nur ein Hinweis.
# Hier nicht direkt verwendet, aber für zukünftige Protokoll-Dateien relevant.
$CsvSeparator = ";"
# EndRegion

# Region: Funktionen
# --------------------
function Test-ADModule
{
    <#
    .SYNOPSIS
        Überprüft, ob das Active Directory Modul geladen ist und versucht, es zu laden.
    #>
    [CmdletBinding()]
    param()

    Write-Verbose "Überprüfe, ob das Active Directory Modul geladen ist..."
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-Error "Das Active Directory Modul ist nicht verfügbar. Bitte stellen Sie sicher, dass die RSAT-Tools installiert sind."
        return $false
    }

    if (-not (Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue)) {
        try {
            Import-Module ActiveDirectory -ErrorAction Stop
            Write-Verbose "Active Directory Modul erfolgreich geladen."
            return $true
        }
        catch {
            Write-Error "Fehler beim Laden des Active Directory Moduls: $($_.Exception.Message)"
            return $false
        }
    }
    Write-Verbose "Active Directory Modul ist bereits geladen."
    return $true
}
# EndRegion

# Region: Skript-Start und Vorbereitung
# -------------------------------------
# Verbose-Ausgabe standardmäßig aktivieren, außer wenn Silent gesetzt ist
if (-not $Silent) {
    $VerbosePreference = "Continue"
} else {
    $VerbosePreference = "SilentlyContinue"
}

Write-Verbose "Skriptausführung gestartet am: $((Get-Date).ToString('G', $CurrentCulture))"
Write-Verbose "Suchmuster: '$SearchPattern'"
if ($OrganizationalUnit) {
    Write-Verbose "Organisationseinheit (OU): '$OrganizationalUnit'"
} else {
    Write-Verbose "Es wurde keine spezifische Organisationseinheit angegeben. Es wird im gesamten AD gesucht."
}
Write-Verbose "Bestätigungsaufforderung pro Benutzer: $($ConfirmAction)"
Write-Verbose "Stille Ausführung (Silent): $($Silent)"

# Modulprüfung
if (-not (Test-ADModule)) {
    Write-Error "Das Skript kann ohne das Active Directory Modul nicht fortgesetzt werden."
    exit 1
}
# EndRegion

# Region: Benutzer suchen
# -----------------------
Write-Verbose "Suche nach Active Directory Benutzern mit dem Namen '$SearchPattern'..."
$UserFilter = "Name -like '$SearchPattern'"

try {
    if ($OrganizationalUnit) {
        $Users = Get-ADUser -Filter $UserFilter -SearchBase $OrganizationalUnit -Properties SamAccountName, PasswordChangeRequired -ErrorAction Stop
        Write-Verbose "Suche in OU '$OrganizationalUnit' abgeschlossen."
    } else {
        $Users = Get-ADUser -Filter $UserFilter -Properties SamAccountName, PasswordChangeRequired -ErrorAction Stop
        Write-Verbose "Globale Suche im Active Directory abgeschlossen."
    }
}
catch {
    Write-Error "Fehler beim Suchen der Benutzer im Active Directory: $($_.Exception.Message)"
    exit 1
}

if (-not $Users) {
    # Korrigierte Zeile: Vereinfachung des Ternäroperators im String
    $LocationText = if ($OrganizationalUnit) { "in der OU '$OrganizationalUnit'" } else { "im gesamten AD" }
    Write-Warning "Keine Benutzer mit dem Suchmuster '$SearchPattern' gefunden $LocationText."
    exit 0
}

Write-Verbose "$($Users.Count) Benutzer gefunden, die dem Suchmuster entsprechen."
# EndRegion

# Region: Benutzer verarbeiten
# ----------------------------
foreach ($User in $Users) {
    Write-Host "" # Leerzeile für bessere Lesbarkeit
    Write-Host "Verarbeite Benutzer: $($User.Name) (sAMAccountName: $($User.SamAccountName))" -ForegroundColor Cyan
    Write-Verbose "Aktueller Status von PasswordChangeRequired für $($User.Name): $($User.PasswordChangeRequired)"

    if ($ConfirmAction) {
        $ConfirmPrompt = "Soll 'PasswordChangeRequired' für Benutzer '$($User.Name)' auf '$false' gesetzt werden?"
        if (-not ($PSCmdlet.ShouldProcess($User.Name, "Set-PasswordChangeRequired - $User.Name"))) {
            Write-Warning "Änderung für $($User.Name) übersprungen (vom Benutzer abgebrochen)."
            continue
        }
        # In PowerShell 5.1 ist ShouldProcess alleine nicht immer ausreichend für direkte Bestätigung
        # wenn der Parameter ConfirmAction manuell gesteuert wird und nicht nur über SupportsShouldProcess.
        # Eine explizite Read-Host Abfrage ist hier oft robuster für interaktive Skripte.
        $UserConfirmed = Read-Host "$ConfirmPrompt (J/N)"
        if ($UserConfirmed -notmatch '^[jJ]$') {
            Write-Warning "Änderung für $($User.Name) übersprungen (vom Benutzer abgelehnt)."
            continue
        }
    } elseif ($PSCmdlet.ShouldProcess($User.Name, "Set-PasswordChangeRequired - $User.Name")) {
        # Wenn ConfirmAction auf $false gesetzt ist, aber ShouldProcess trotzdem ausgeführt wird
        # (z.B. durch -Confirm bei Skriptaufruf), dann einfach fortfahren
        Write-Verbose "Bestätigungsabfrage ist deaktiviert oder wurde akzeptiert. Setze 'PasswordChangeRequired' für $($User.Name)."
    } else {
        Write-Warning "Änderung für $($User.Name) übersprungen (ShouldProcess abgelehnt, z.B. durch -WhatIf)."
        continue
    }

    try {
        Set-ADUser -Identity $User.SamAccountName -PasswordChangeRequired $false -ErrorAction Stop
        Write-Host "ERFOLG: 'PasswordChangeRequired' für Benutzer '$($User.Name)' auf '$false' gesetzt." -ForegroundColor Green

        # Überprüfung des Erfolgs
        $UpdatedUser = Get-ADUser -Identity $User.SamAccountName -Properties PasswordChangeRequired -ErrorAction Stop
        if ($UpdatedUser.PasswordChangeRequired -eq $false) {
            Write-Verbose "Verifizierung erfolgreich: '$($User.Name)' hat nun 'PasswordChangeRequired' = '$false'."
        } else {
            Write-Warning "Verifizierung fehlgeschlagen: 'PasswordChangeRequired' für '$($User.Name)' ist immer noch '$($UpdatedUser.PasswordChangeRequired)'."
        }
    }
    catch {
        Write-Error "FEHLER: Konnte 'PasswordChangeRequired' für Benutzer '$($User.Name)' nicht ändern: $($_.Exception.Message)"
    }
}
# EndRegion

Write-Verbose "Skriptausführung beendet am: $((Get-Date).ToString('G', $CurrentCulture))"
Write-Host "`nSkriptausführung abgeschlossen." -ForegroundColor Yellow
