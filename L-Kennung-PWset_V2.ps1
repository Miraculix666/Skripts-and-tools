<#
.SYNOPSIS
AD-Benutzerverwaltungsskript für spezielle OUs

.DESCRIPTION
Dieses Skript führt folgende Aufgaben durch:
1. Dynamische Suche nach OUs 81 und 82
2. Ermittlung aller Benutzer mit Namen L110* oder L114*
3. Passwortzurücksetzung mit Aktivierung
4. Deutschsprachige Berichterstellung

.PARAMETER TargetOUNames
Kommagetrennte Liste der OU-Namen (Standard: 81,82)

.PARAMETER ReportPath
Ausgabepfad für Berichtsdatei (Standard: C:\Daten\Benutzerbericht.csv)

.PARAMETER Force
Unterdrückt Bestätigungsaufforderungen

.EXAMPLE
.\Skript.ps1 -TargetOUNames 81,82 -Verbose
#>

#Requires -Version 5.1
#Requires -Modules ActiveDirectory

[CmdletBinding()]
param(
    [Parameter()]
    [string[]]$TargetOUNames = @('81', '82'),
    
    [string]$ReportPath = "C:\Daten\Benutzerbericht.csv",
    
    [switch]$Force
)

begin {
    # Deutsche Lokalisierung
    $culture = [System.Globalization.CultureInfo]::GetCultureInfo('de-DE')
    [System.Threading.Thread]::CurrentThread.CurrentCulture = $culture
    [System.Threading.Thread]::CurrentThread.CurrentUICulture = $culture

    # Initialisierung
    $ErrorActionPreference = 'Stop'
    $startTime = Get-Date
    $stats = [PSCustomObject]@{
        TotalUsers  = 0
        Success     = 0
        Errors      = 0
        ErrorList   = [System.Collections.Generic.List[object]]::new()
    }
}

process {
    try {
        # Dynamische OU-Suche -------------------------------------------------
        Write-Verbose "Starte OU-Suche"
        $domain = Get-ADDomain
        Write-Verbose "Domain: $($domain.DNSRoot)"
        
        $targetOUs = foreach ($name in $TargetOUNames) {
            Write-Verbose "Suche OU: $name"
            Get-ADOrganizationalUnit -Filter "Name -like '*$name*'" -SearchBase $domain.DistinguishedName -SearchScope Subtree
        }

        if (-not $targetOUs) {
            throw "Keine OUs gefunden für: $($TargetOUNames -join ', ')"
        }

        Write-Verbose "Gefundene OUs:`n$($targetOUs | Format-Table Name, DistinguishedName -AutoSize | Out-String)"

        # Benutzersuche -------------------------------------------------------
        $userFilter = "Name -like 'L11[04]*'"
        Write-Verbose "Verwende Benutzerfilter: $userFilter"
        
        $users = foreach ($ou in $targetOUs) {
            Write-Verbose "Durchsuche OU: $($ou.DistinguishedName)"
            Get-ADUser -Filter $userFilter -SearchBase $ou.DistinguishedName -Properties *
        }

        if (-not $users) {
            Write-Host "Keine passenden Benutzer gefunden." -ForegroundColor Yellow
            return
        }

        $stats.TotalUsers = $users.Count
        Write-Verbose "Gefundene Benutzer: $($stats.TotalUsers)"

        # CSV-Export ----------------------------------------------------------
        $users | Select-Object @{n='Benutzername';e={$_.Name}},
                              @{n='Anmeldename';e={$_.SamAccountName}},
                              @{n='Aktiviert';e={$_.Enabled}},
                              @{n='Letzte Anmeldung';e={$_.LastLogonDate}} |
                 Export-Csv -Path $ReportPath -Delimiter ";" -Encoding UTF8 -NoTypeInformation
        
        Write-Host "Bericht erstellt: $ReportPath" -ForegroundColor Green

        # Bestätigung ---------------------------------------------------------
        if (-not $Force) {
            $confirmation = Read-Host "Möchten Sie fortfahren? (J/N)"
            if ($confirmation -notin @('J','j')) { return }
        }

        # Passwortverwaltung --------------------------------------------------
        $securePass = Read-Host "Neues Passwort eingeben" -AsSecureString
        $plainPass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePass)
        )

        foreach ($user in $users) {
            try {
                Write-Verbose "Verarbeite Benutzer: $($user.SamAccountName)"
                
                # Net User-Kommando
                $output = net user $user.SamAccountName $plainPass /DOMAIN /ACTIVE:YES /PASSWORDCHG:NO 2>&1
                if ($LASTEXITCODE -ne 0) {
                    throw "Net User Fehler ($LASTEXITCODE): $output"
                }

                # AD-Einstellungen
                Set-ADUser -Identity $user -PasswordNeverExpires $true `
                    -CannotChangePassword $true `
                    -Replace @{lastLogonTimestamp = [DateTime]::Now.ToFileTime()} `
                    -ErrorAction Stop

                $stats.Success++
                Write-Host "$($user.SamAccountName): Erfolgreich aktualisiert" -ForegroundColor Green
            }
            catch {
                $stats.Errors++
                $errorMsg = $_.Exception.Message
                $stats.ErrorList.Add([PSCustomObject]@{
                    Benutzer = $user.SamAccountName
                    Fehler = $errorMsg
                })
                Write-Host "Fehler bei $($user.SamAccountName): $errorMsg" -ForegroundColor Red
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
    $duration = (Get-Date) - $startTime
    Write-Host @"
    
    ===== ZUSAMMENFASSUNG =====
    Verarbeitete Benutzer: $($stats.TotalUsers)
    Erfolgreich:           $($stats.Success)
    Fehler:                $($stats.Errors)
    Dauer:                 $($duration.ToString('hh\:mm\:ss'))
    
    Fehlerdetails:
    $($stats.ErrorList | Format-Table -AutoSize | Out-String)
"@

    # Ergebnisrückgabe
    if ($stats.Errors -gt 0) { exit 2 }
}
