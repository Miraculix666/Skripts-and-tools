<#
.SYNOPSIS
AD-Benutzerverwaltungsskript mit deutscher Lokalisierung

.DESCRIPTION
Dieses Skript führt folgende Aufgaben durch:
1. Findet Benutzer in spezifizierten OUs
2. Exportiert Berichte im deutschen CSV-Format
3. Setzt Passwörter und Kontoeigenschaften
4. Bietet detaillierte Protokollierung

.VERSION
2.1.20250407

.PARAMETER TargetOUNames
Kommagetrennte Liste der OU-Namen (Standard: 81,82)

.PARAMETER ReportPath
Ausgabepfad für Berichtsdatei (Standard: C:\Daten\Benutzerbericht.csv)

.PARAMETER Verbose
Aktiviert detaillierte Protokollierung

.EXAMPLE
.\Skript.ps1 -TargetOUNames 81,82 -Verbose
#>

#Requires -Version 5.1
#Requires -Modules ActiveDirectory

[CmdletBinding()]
param(
    [Parameter()]
    [string[]]$TargetOUNames = @('81', '82'),
    
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
        # OU-Suche
        Write-Verbose "Starte OU-Suche für: $($TargetOUNames -join ', ')"
        $domain = Get-ADDomain
        $searchBase = $domain.DistinguishedName
        
        $ous = foreach ($name in $TargetOUNames) {
            Get-ADOrganizationalUnit -Filter "Name -eq '$name'" -SearchBase $searchBase -SearchScope Subtree
        }

        if (-not $ous) {
            throw "Keine OUs gefunden für: $($TargetOUNames -join ', ')"
        }

        # Benutzerabfrage
        $userFilter = "Name -like 'L11[04]*'"
        Write-Verbose "Verwende Filter: $userFilter"
        
        $users = foreach ($ou in $ous) {
            Get-ADUser -Filter $userFilter -SearchBase $ou.DistinguishedName -Properties *
        }

        if (-not $users) {
            Write-Host "Keine passenden Benutzer gefunden." -ForegroundColor Yellow
            return
        }

        $stats.TotalUsers = $users.Count

        # CSV-Export
        $users | Select-Object @{n='Benutzername';e={$_.Name}},
                              @{n='Anmeldename';e={$_.SamAccountName}},
                              @{n='Aktiviert';e={$_.Enabled}},
                              @{n='Letzte Anmeldung';e={$_.LastLogonDate}} |
                 Export-Csv -Path $ReportPath -Delimiter ";" -Encoding UTF8 -NoTypeInformation
        
        Write-Verbose "Bericht erstellt: $ReportPath"

        # Bestätigung
        if (-not $Force) {
            $confirmation = Read-Host "Möchten Sie fortfahren? (J/N)"
            if ($confirmation -notin @('J','j')) { return }
        }

        # Passwortverwaltung
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
                Set-ADUser -Identity $user -PasswordNeverExpires $true -CannotChangePassword $true
                Set-ADUser -Identity $user -Replace @{lastLogonTimestamp = [DateTime]::Now.ToFileTime()}

                $stats.Success++
            }
            catch {
                $stats.Errors++
                $stats.ErrorList.Add([PSCustomObject]@{
                    Benutzer = $user.SamAccountName
                    Fehler = $_.Exception.Message
                })
                Write-Warning "Fehler bei $($user.SamAccountName): $($_.Exception.Message)"
            }
        }
    }
    catch {
        Write-Error "Kritischer Fehler: $($_.Exception.Message)"
        exit 1
    }
    finally {
        # Bereinigung
        Remove-Variable plainPass -ErrorAction SilentlyContinue
    }
}

end {
    # Zusammenfassung
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
}
