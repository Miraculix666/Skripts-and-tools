<#
.SYNOPSIS
    Erstellt AD-Benutzer aus CSV-Datei als Kopie eines Vorlagenbenutzers
    
.DESCRIPTION
    Dieses Skript erstellt neue AD-Benutzer basierend auf:
    - Einer CSV-Datei mit Benutzerdaten
    - Einem Vorlagenbenutzer für Standardattribute
    - Passwort-Compliance-Prüfungen
    - Automatischer Gruppenübernahme
    - Optionale Post-Erstellungsverifikation

.PARAMETER TemplateUser
    SAMAccountName des Vorlagenbenutzers

.PARAMETER CsvPath
    Pfad zur CSV-Datei mit folgenden Spalten:
    NewUserName,NewUserPassword,FirstName,LastName,Department

.PARAMETER Verify
    Aktiviert die Überprüfung nach der Erstellung

.EXAMPLE
    .\Create-ADUsers.ps1 -TemplateUser "VorlagenUser" -CsvPath "users.csv" -Verify

.NOTES
    Version: 3.0 | Autor: IT-Team | Letzte Änderung: 2025-02-04
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [string]$TemplateUser,

    [Parameter(Mandatory = $true)]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [string]$CsvPath,

    [switch]$Verify
)

begin {
    #region Hilfsfunktionen
    
    function Test-PasswordComplexity {
        param([string]$Password)
        $complexityRules = @(
            { $Password.Length -ge 12 },
            { $Password -match '[A-Z]' },
            { $Password -match '[a-z]' },
            { $Password -match '\d' },
            { $Password -match '[\W_]' }
        )
        return ($complexityRules | Where-Object { -not (& $_) }).Count -eq 0
    }

    function Copy-ADGroupMembership {
        param(
            [string]$SourceUser,
            [string]$TargetUser
        )
        try {
            $groups = Get-ADUser -Identity $SourceUser -Properties MemberOf | 
                     Select-Object -ExpandProperty MemberOf
            foreach ($group in $groups) {
                Add-ADGroupMember -Identity $group -Members $TargetUser -ErrorAction Stop
            }
            Write-Host "Gruppen kopiert: $($groups.Count)" -ForegroundColor Green
        }
        catch {
            Write-Host "Fehler beim Gruppenkopieren: $_" -ForegroundColor Red
            throw
        }
    }

    function New-ADUserFromTemplate {
        param(
            [string]$TemplateSamAccountName,
            [string]$NewSamAccountName,
            [string]$Password,
            [hashtable]$AdditionalAttributes
        )
        try {
            # Vorlagenbenutzer laden
            $template = Get-ADUser -Identity $TemplateSamAccountName -Properties *

            # Benutzerparameter erstellen
            $userParams = @{
                Instance              = $template
                SamAccountName       = $NewSamAccountName
                UserPrincipalName    = "$NewSamAccountName@$((Get-ADDomain).DNSRoot)"
                Name                 = $AdditionalAttributes.DisplayName
                GivenName            = $AdditionalAttributes.FirstName
                Surname              = $AdditionalAttributes.LastName
                Department           = $AdditionalAttributes.Department
                AccountPassword      = ConvertTo-SecureString $Password -AsPlainText -Force
                Enabled              = $true
                ChangePasswordAtLogon = $true
                PassThru             = $true
            }

            # Benutzer erstellen
            $newUser = New-ADUser @userParams -ErrorAction Stop
            Write-Host "Benutzer $NewSamAccountName erstellt" -ForegroundColor Green

            # Gruppen kopieren
            Copy-ADGroupMembership -SourceUser $TemplateSamAccountName -TargetUser $NewSamAccountName

            return $newUser
        }
        catch {
            Write-Host "Fehler bei $NewSamAccountName: $_" -ForegroundColor Red
            return $null
        }
    }
    #endregion
}

process {
    try {
        # CSV-Daten importieren und validieren
        $users = Import-Csv -Path $CsvPath -Encoding UTF8
        $requiredColumns = @('NewUserName','NewUserPassword','FirstName','LastName','Department')
        $missingColumns = $requiredColumns | Where-Object { $_ -notin $users[0].PSObject.Properties.Name }

        if ($missingColumns) {
            throw "Fehlende Spalten in CSV: $($missingColumns -join ', ')"
        }

        # Hauptverarbeitung
        $createdUsers = @()
        foreach ($user in $users) {
            # Passwortvalidierung
            if (-not (Test-PasswordComplexity -Password $user.NewUserPassword)) {
                Write-Host "Passwort für $($user.NewUserName) entspricht nicht den Richtlinien" -ForegroundColor Red
                continue
            }

            # Benutzererstellung
            $additionalAttributes = @{
                FirstName   = $user.FirstName
                LastName    = $user.LastName
                Department  = $user.Department
                DisplayName = "$($user.FirstName) $($user.LastName)"
            }

            $newUser = New-ADUserFromTemplate -TemplateSamAccountName $TemplateUser `
                                             -NewSamAccountName $user.NewUserName `
                                             -Password $user.NewUserPassword `
                                             -AdditionalAttributes $additionalAttributes
            
            if ($newUser) {
                $createdUsers += $newUser
            }
        }

        # Verifikation
        if ($Verify -and $createdUsers.Count -gt 0) {
            Write-Host "`nVerifikation der erstellten Benutzer:" -ForegroundColor Cyan
            $createdUsers | ForEach-Object {
                $user = Get-ADUser -Identity $_.SamAccountName -Properties Department,MemberOf
                Write-Host "`nBenutzer: $($user.SamAccountName)"
                Write-Host "Abteilung: $($user.Department)"
                Write-Host "Gruppen: $($user.MemberOf.Count)"
            }
        }
    }
    catch {
        Write-Host "KRITISCHER FEHLER: $_" -ForegroundColor Red
        exit 1
    }
}

end {
    Write-Host "`nProzess abgeschlossen. Erstellte Benutzer: $($createdUsers.Count)" -ForegroundColor Cyan
}
