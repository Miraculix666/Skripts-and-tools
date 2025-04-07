<#
.SYNOPSIS
AD-Benutzerverwaltung mit erweitertem Logging und deutscher Lokalisierung

.DESCRIPTION
Dieses Skript führt folgende Aufgaben durch:
1. Sucht Benutzer in OUs 81/82 mit Namensmustern L110* oder L114*
2. Zeigt gefundene Benutzer an
3. Protokolliert alle Operationen in CSV
4. Setzt Passwörter zurück und aktualisiert Kontoeigenschaften
5. Deutsche Lokalisierung und Formatierung

.VERSION
3.0.20250407
#>

#Requires -Version 5.1
#Requires -Modules ActiveDirectory

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter()]
    [string[]]$OUNames = @('81', '82'),

    [Parameter()]
    [string]$ReportPath = "C:\Daten\Benutzerbericht.csv",

    [Parameter()]
    [switch]$Force
)

begin {
    # Deutsche Lokalisierung
    $culture = [System.Globalization.CultureInfo]::GetCultureInfo('de-DE')
    [System.Threading.Thread]::CurrentThread.CurrentCulture = $culture
    [System.Threading.Thread]::CurrentThread.CurrentUICulture = $culture

    # Initialisierungen
    $ErrorActionPreference = 'Stop'
    $global:OperationResults = [System.Collections.Generic.List[object]]::new()
    $PasswordSetDate = Get-Date

    # Protokollierungsfunktion
    function Write-Header {
        param([string]$Title)
        Write-Host "`n$(('=' * 80))" -ForegroundColor Cyan
        Write-Host " $Title " -ForegroundColor Cyan
        Write-Host "$(('=' * 80))`n" -ForegroundColor Cyan
    }
}

process {
    try {
        # 1. OU-SUCHE ---------------------------------------------------------
        Write-Header -Title "SCHRITT 1: ORGANISATIONSEINHEITEN-SUCHE"
        Write-Host "Suche nach OUs: $($OUNames -join ', ')" -ForegroundColor Yellow
        
        $domain = Get-ADDomain
        $targetOUs = foreach ($name in $OUNames) {
            Get-ADOrganizationalUnit -Filter "Name -eq '$name'" `
                -SearchBase $domain.DistinguishedName `
                -SearchScope Subtree
        }

        if (-not $targetOUs) {
            throw "Keine OUs gefunden für: $($OUNames -join ', ')"
        }

        # 2. BENUTZERSUCHE ----------------------------------------------------
        Write-Header -Title "SCHRITT 2: BENUTZERRECHE"
        $users = foreach ($ou in $targetOUs) {
            Get-ADUser -LDAPFilter "(|(sAMAccountName=L110*)(sAMAccountName=L114*))" `
                -SearchBase $ou.DistinguishedName `
                -Properties * `
                -SearchScope Subtree
        }

        if (-not $users) {
            Write-Host "Keine passenden Benutzer gefunden." -ForegroundColor Yellow
            return
        }

        # 3. BENUTZERANZEIGE --------------------------------------------------
        Write-Header -Title "GEFUNDENE BENUTZER"
        $users | Format-Table @{l='Benutzername';e={$_.Name}},
                            @{l='Aktiviert';e={if($_.Enabled){'Ja'}else{'Nein'}}},
                            @{l='Letzte Anmeldung';e={$_.LastLogonDate}} -AutoSize

        # 4. PASSWORTHINWEIS --------------------------------------------------
        Write-Header -Title "PASSWORTINFORMATION"
        Write-Host "Geplantes Passwort-Änderungsdatum: $($PasswordSetDate.ToString('dd.MM.yyyy HH:mm:ss'))" -ForegroundColor Magenta

        # 5. BESTÄTIGUNG ------------------------------------------------------
        Write-Header -Title "BESTÄTIGUNG"
        if (-not $Force) {
            $confirmation = Read-Host "Möchten Sie die Änderungen durchführen? (J/N)"
            if ($confirmation -notin @('J','j')) { return }
        }

        # 6. PASSWORTVERWALTUNG -----------------------------------------------
        Write-Header -Title "PASSWORTZURÜCKSETZUNG"
        $securePass = Read-Host "Neues Passwort eingeben" -AsSecureString
        $plainPass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePass)
        )

        # 7. BENUTZERVERARBEITUNG ---------------------------------------------
        foreach ($user in $users) {
            $resultTemplate = [PSCustomObject]@{
                Benutzername     = $user.Name
                Anmeldename      = $user.SamAccountName
                Aktiviert        = $user.Enabled
                LetzteAnmeldung  = $user.LastLogonDate
                OU_Pfad          = $user.DistinguishedName
                OperationStatus  = "Nicht durchgeführt"
                FehlerCode       = $null
                FehlerMeldung    = $null
                PasswortDatum    = $null
            }

            try {
                Write-Host "`nVerarbeite Benutzer: $($user.SamAccountName)" -ForegroundColor Cyan

                if ($PSCmdlet.ShouldProcess($user.SamAccountName, "Passwort zurücksetzen")) {
                    # Passwortreset
                    $output = net user $user.SamAccountName $plainPass /DOMAIN /ACTIVE:YES 2>&1
                    if ($LASTEXITCODE -ne 0) {
                        throw "NetUser-Fehler ($LASTEXITCODE): $output"
                    }

                    # AD-Einstellungen
                    Set-ADUser -Identity $user -PasswordNeverExpires $true `
                        -CannotChangePassword $true `
                        -Replace @{lastLogonTimestamp = [DateTime]::Now.ToFileTime()}

                    # Erfolgsmeldung
                    $resultTemplate.OperationStatus = "Erfolgreich"
                    $resultTemplate.PasswortDatum = Get-Date
                    Write-Host "Erfolgreich aktualisiert" -ForegroundColor Green
                }
            }
            catch {
                $resultTemplate.OperationStatus = "Fehlgeschlagen"
                $resultTemplate.FehlerCode = $LASTEXITCODE
                $resultTemplate.FehlerMeldung = $_.Exception.Message
                $resultTemplate.PasswortDatum = $null
                
                Write-Host "Fehler: $($_.Exception.Message)" -ForegroundColor Red
            }
            finally {
                $global:OperationResults.Add($resultTemplate)
            }
        }
    }
    catch {
        Write-Host "KRITISCHER FEHLER: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
    finally {
        # 8. CSV-EXPORT -------------------------------------------------------
        Write-Header -Title "CSV-EXPORT"
        $global:OperationResults | Export-Csv -Path $ReportPath -Delimiter ";" -Encoding UTF8 -NoTypeInformation
        Write-Host "Bericht erstellt: $ReportPath" -ForegroundColor Green

        # Bereinigung
        Remove-Variable plainPass -ErrorAction SilentlyContinue
    }
}

end {
    # 9. ZUSAMMENFASSUNG ------------------------------------------------------
    Write-Header -Title "ZUSAMMENFASSUNG"
    $successCount = ($global:OperationResults | Where-Object { $_.OperationStatus -eq "Erfolgreich" }).Count
    $errorCount = ($global:OperationResults | Where-Object { $_.OperationStatus -eq "Fehlgeschlagen" }).Count

    Write-Host @"
    Verarbeitete Benutzer: $($global:OperationResults.Count)
    Erfolgreich:          $successCount
    Fehlgeschlagen:       $errorCount

    Fehlerübersicht:
    $($global:OperationResults | Where-Object { $_.FehlerCode } | Format-Table Anmeldename, FehlerCode, FehlerMeldung -AutoSize | Out-String)
"@

    if ($errorCount -gt 0) { exit 2 }
}
