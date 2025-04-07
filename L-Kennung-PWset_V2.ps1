<#
.SYNOPSIS
AD-Benutzerverwaltung mit Passwortreset und deutscher Lokalisierung

.DESCRIPTION
Dieses Skript führt folgende Aufgaben durch:
1. Findet Benutzer in OUs 81/82 mit Namensmustern L110* oder L114*
2. Exportiert Ergebnisse in CSV mit deutschen Formatierung
3. Passwortreset mit Bestätigungsaufforderung
4. Detailliertes Logging und Fehlerbehandlung

.VERSION
2.2.20250407
#>

#Requires -Version 5.1
#Requires -Modules ActiveDirectory

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [string[]]$OUNames = @('81', '82'),

    [Parameter(Mandatory = $false)]
    [string]$ReportPath = "C:\Daten\Benutzerbericht.csv",

    [Parameter(Mandatory = $false)]
    [switch]$Force
)

begin {
    # Deutsche Lokalisierung
    $culture = [System.Globalization.CultureInfo]::GetCultureInfo('de-DE')
    [System.Threading.Thread]::CurrentThread.CurrentCulture = $culture
    [System.Threading.Thread]::CurrentThread.CurrentUICulture = $culture

    # Initialisierungen
    $ErrorActionPreference = 'Stop'
    $stats = [PSCustomObject]@{
        GefundeneBenutzer = 0
        Erfolgreich       = 0
        Fehler            = 0
        Fehlerliste       = [System.Collections.Generic.List[string]]::new()
    }

    # Protokollierungsfunktion
    function Write-Log {
        param([string]$Message, [string]$Level = "INFO")
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Write-Verbose "[$timestamp][$Level] $Message"
    }
}

process {
    try {
        # 1. OU-Suche ---------------------------------------------------------
        Write-Log "Starte OU-Suche für: $($OUNames -join ', ')"
        $domain = Get-ADDomain
        $targetOUs = foreach ($name in $OUNames) {
            Get-ADOrganizationalUnit -Filter "Name -eq '$name'" `
                -SearchBase $domain.DistinguishedName `
                -SearchScope Subtree
        }

        if (-not $targetOUs) {
            throw "Ziel-OUs nicht gefunden!"
        }

        # 2. Benutzersuche ----------------------------------------------------
        Write-Log "Starte Benutzersuche"
        $users = foreach ($ou in $targetOUs) {
            Get-ADUser -LDAPFilter "(|(sAMAccountName=L110*)(sAMAccountName=L114*))" `
                -SearchBase $ou.DistinguishedName `
                -Properties Enabled, LastLogonDate, PasswordNeverExpires `
                -SearchScope Subtree
        }

        if (-not $users) {
            Write-Host "Keine passenden Benutzer gefunden." -ForegroundColor Yellow
            return
        }

        $stats.GefundeneBenutzer = $users.Count

        # 3. Bildschirmausgabe ------------------------------------------------
        Write-Host "`nGefundene Benutzer ($($stats.GefundeneBenutzer)):"
        $users | Format-Table @{l='Benutzername';e={$_.Name}}, 
                            @{l='Aktiviert';e={if($_.Enabled){'Ja'}else{'Nein'}}}, 
                            @{l='Letzte Anmeldung';e={$_.LastLogonDate}} -AutoSize

        # 4. CSV-Export -------------------------------------------------------
        $users | Select-Object @{n='Benutzername';e={$_.Name}},
                              @{n='Anmeldename';e={$_.SamAccountName}},
                              @{n='Aktiviert';e={$_.Enabled}},
                              @{n='Letzte Anmeldung';e={$_.LastLogonDate}},
                              @{n='OU-Pfad';e={$_.DistinguishedName}} |
                 Export-Csv -Path $ReportPath -Delimiter ";" -Encoding UTF8 -NoTypeInformation
        
        Write-Host "`nBericht erstellt: $ReportPath" -ForegroundColor Green

        # 5. Bestätigung ------------------------------------------------------
        if (-not $Force) {
            $confirmation = Read-Host "`nMöchten Sie die Änderungen durchführen? (J/N)"
            if ($confirmation -notin @('J','j')) { return }
        }

        # 6. Passwortverwaltung -----------------------------------------------
        $securePass = Read-Host "Neues Passwort eingeben" -AsSecureString
        $plainPass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePass)
        )

        # 7. Benutzerverarbeitung ---------------------------------------------
        foreach ($user in $users) {
            try {
                Write-Log "Verarbeite $($user.SamAccountName)"
                
                if ($PSCmdlet.ShouldProcess($user.SamAccountName, "Passwort zurücksetzen")) {
                    # Passwortreset mit net user
                    $output = net user $user.SamAccountName $plainPass /DOMAIN /ACTIVE:YES 2>&1
                    if ($LASTEXITCODE -ne 0) {
                        throw "Net User Fehler ($LASTEXITCODE): $output"
                    }

                    # AD-Einstellungen
                    Set-ADUser -Identity $user -PasswordNeverExpires $true `
                        -CannotChangePassword $true `
                        -Replace @{lastLogonTimestamp = [DateTime]::Now.ToFileTime()}

                    $stats.Erfolgreich++
                    Write-Host "$($user.SamAccountName) erfolgreich aktualisiert" -ForegroundColor Green
                }
            }
            catch {
                $stats.Fehler++
                $errorMsg = "Fehler bei $($user.SamAccountName): $($_.Exception.Message)"
                $stats.Fehlerliste.Add($errorMsg)
                Write-Host $errorMsg -ForegroundColor Red
            }
        }
    }
    catch {
        Write-Host "KRITISCHER FEHLER: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
    finally {
        # Bereinigung
        Remove-Variable plainPass -ErrorAction SilentlyContinue
    }
}

end {
    # Zusammenfassung ---------------------------------------------------------
    Write-Host @"
`n========== ZUSAMMENFASSUNG ==========
Gefundene Benutzer:   $($stats.GefundeneBenutzer)
Erfolgreich:          $($stats.Erfolgreich)
Fehler:               $($stats.Fehler)
`nFehlerdetails:
$($stats.Fehlerliste -join "`n")
`n
"@

    if ($stats.Fehler -gt 0) { exit 2 }
}
