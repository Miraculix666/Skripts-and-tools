<#
.SYNOPSIS
AD-Benutzerverwaltung mit Passwortzurücksetzung und deutscher Lokalisierung

.DESCRIPTION
Dieses Skript führt folgende Aufgaben durch:
1. Findet Benutzer in den OUs 81/82 mit Namensmustern L110* oder L114*
2. Exportiert Ergebnisse in eine CSV-Datei
3. Setzt Passwörter zurück und konfiguriert Kontoeigenschaften
4. Aktualisiert Anmeldezeitstempel

.NOTES
Version: 1.5
Erstellt am: 2024-03-15
#>

#Requires -Version 5.1
#Requires -Modules ActiveDirectory

[CmdletBinding()]
param(
    [Parameter()]
    [string[]]$OUNames = @('81', '82'),

    [Parameter()]
    [string]$ReportPath = "C:\Daten\Benutzerbericht.csv",

    [Parameter()]
    [securestring]$NewPassword,

    [switch]$Force
)

begin {
    # Deutsche Lokalisierung
    $culture = [System.Globalization.CultureInfo]::GetCultureInfo('de-DE')
    [System.Threading.Thread]::CurrentThread.CurrentCulture = $culture
    [System.Threading.Thread]::CurrentThread.CurrentUICulture = $culture

    # Initialisierung
    $ErrorActionPreference = 'Stop'
    $stats = [PSCustomObject]@{
        GefundeneBenutzer = 0
        Erfolgreich       = 0
        Fehler            = 0
        Fehlerliste       = [System.Collections.Generic.List[string]]::new()
    }
}

process {
    try {
        # 1. OU-Suche ---------------------------------------------------------
        Write-Verbose "Starte OU-Suche"
        $domain = Get-ADDomain
        $targetOUs = foreach ($name in $OUNames) {
            Get-ADOrganizationalUnit -Filter "Name -eq '$name'" `
                -SearchBase $domain.DistinguishedName `
                -SearchScope Subtree
        }

        if (-not $targetOUs) {
            throw "Ziel-OUs ($($OUNames -join ', ')) nicht gefunden!"
        }

        # 2. Benutzersuche ----------------------------------------------------
        $users = foreach ($ou in $targetOUs) {
            Get-ADUser -LDAPFilter "(|(sAMAccountName=L110*)(sAMAccountName=L114*))" `
                -SearchBase $ou.DistinguishedName `
                -Properties Enabled, LastLogonDate `
                -SearchScope Subtree
        }

        if (-not $users) {
            Write-Host "Keine passenden Benutzer gefunden." -ForegroundColor Yellow
            return
        }

        $stats.GefundeneBenutzer = $users.Count

        # 3. CSV-Export -------------------------------------------------------
        $users | Select-Object @{n='Benutzername';e={$_.Name}},
                              @{n='Anmeldename';e={$_.SamAccountName}},
                              @{n='Aktiviert';e={$_.Enabled}},
                              @{n='Letzte Anmeldung';e={$_.LastLogonDate}} |
                 Export-Csv -Path $ReportPath -Delimiter ";" -Encoding UTF8 -NoTypeInformation
        
        Write-Host "Bericht erstellt: $ReportPath" -ForegroundColor Green

        # 4. Bestätigung ------------------------------------------------------
        if (-not $Force) {
            $confirmation = Read-Host "Möchten Sie die Änderungen durchführen? (J/N)"
            if ($confirmation -notin @('J','j')) { return }
        }

        # 5. Passwortverwaltung -----------------------------------------------
        if (-not $NewPassword) {
            $NewPassword = Read-Host "Neues Passwort eingeben" -AsSecureString
        }

        $plainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($NewPassword)
        )

        # 6. Benutzerverarbeitung ---------------------------------------------
        foreach ($user in $users) {
            try {
                Write-Verbose "Verarbeite $($user.SamAccountName)"

                # Passwortzurücksetzung mit net user
                $output = net user $user.SamAccountName $plainPassword /DOMAIN /ACTIVE:YES 2>&1
                if ($LASTEXITCODE -ne 0) {
                    throw "Net User Fehler: $output"
                }

                # AD-Einstellungen
                Set-ADUser -Identity $user -PasswordNeverExpires $true `
                    -CannotChangePassword $true `
                    -Replace @{lastLogonTimestamp = [DateTime]::Now.ToFileTime()}

                $stats.Erfolgreich++
                Write-Host "$($user.SamAccountName) erfolgreich aktualisiert" -ForegroundColor Green
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
        Remove-Variable plainPassword -ErrorAction SilentlyContinue
    }
}

end {
    # Zusammenfassung ---------------------------------------------------------
    Write-Host @"
    
    ===== ZUSAMMENFASSUNG =====
    Gefundene Benutzer:   $($stats.GefundeneBenutzer)
    Erfolgreich:          $($stats.Erfolgreich)
    Fehler:               $($stats.Fehler)
    
    Fehlerdetails:
    $($stats.Fehlerliste -join "`n")
"@

    if ($stats.Fehler -gt 0) { exit 2 }
}
