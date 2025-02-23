# Funktion zum Erstellen von Benutzern aus CSV
function Import-ADUsersFromCSV {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CsvPath,
        [Parameter(Mandatory = $false)]
        [SecureString]$DefaultPassword = (ConvertTo-SecureString "Willkommen2024!" -AsPlainText -Force)
    )

    # Überprüfe, ob die CSV-Datei existiert
    if (-not (Test-Path $CsvPath)) {
        Write-CustomLog "CSV-Datei nicht gefunden: $CsvPath" -Level "FEHLER"
        return
    }

    Write-CustomLog "Importiere Benutzer aus CSV: $CsvPath" -Level "INFO"

    try {
        # Lese CSV-Datei
        $users = Import-Csv -Path $CsvPath -Delimiter ";" -Encoding UTF8

        foreach ($user in $users) {
            # Bestimme OU basierend auf CSV oder Standard-OU
            $ou = if ($user.OU) { $user.OU } else { $DefaultOU }

            if ($PSCmdlet.ShouldProcess($user.Name, "Benutzer aus CSV erstellen")) {
                $params = @{
                    SamAccountName = $user.Benutzername
                    UserPrincipalName = $user.EMail
                    Name = ($user.Vorname + " " + $user.Nachname)
                    OU = $ou
                    Groups = ($user.Gruppen -split ';')
                    Password = $DefaultPassword
                }

                # Optional: Füge zusätzliche Attribute hinzu, wenn sie im CSV vorhanden sind
                if ($user.Vorname) { $params.GivenName = $user.Vorname }
                if ($user.Nachname) { $params.Surname = $user.Nachname }
                if ($user.Abteilung) { $params.Department = $user.Abteilung }
                if ($user.Position) { $params.Title = $user.Position }
                if ($user.Vorgesetzter) { $params.Manager = $user.Vorgesetzter }
                if ($user.Büro) { $params.Office = $user.Büro }
                if ($user.Telefon) { $params.OfficePhone = $user.Telefon }
                if ($user.Firma) { $params.Company = $user.Firma }
                if ($user.Beschreibung) { $params.Description = $user.Beschreibung }

                New-CustomADUser @params
            }
        }
    } catch {
        Write-CustomLog "Fehler beim Import aus CSV: $_" -Level "FEHLER"
    }
}

# Funktion zum Erstellen eines neuen AD-Benutzers
function New-CustomADUser {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SamAccountName,
        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName,
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $false)]
        [string]$GivenName,
        [Parameter(Mandatory = $false)]
        [string]$Surname,
        [Parameter(Mandatory = $false)]
        [string]$Department,
        [Parameter(Mandatory = $false)]
        [string]$Title,
        [Parameter(Mandatory = $false)]
        [string]$Manager,
        [Parameter(Mandatory = $false)]
        [string]$Office,
        [Parameter(Mandatory = $false)]
        [string]$OfficePhone,
        [Parameter(Mandatory = $false)]
        [string]$Company,
        [Parameter(Mandatory = $false)]
        [string]$Description,
        [Parameter(Mandatory = $true)]
        [string]$OU,
        [string[]]$Groups,
        [SecureString]$Password
    )

    Write-Verbose "Erstelle neuen Benutzer: $Name in OU: $OU"

    try {
        if ($PSCmdlet.ShouldProcess($Name, "Benutzer erstellen")) {
            # Überprüfe, ob die OU existiert
            if (-not (Get-ADOrganizationalUnit -Filter { DistinguishedName -eq $OU })) {
                throw "Die angegebene OU existiert nicht: $OU"
            }

            $userParams = @{
                Name = $Name
                SamAccountName = $SamAccountName
                UserPrincipalName = $UserPrincipalName
                Path = $OU
                AccountPassword = $Password
                Enabled = $true
                ChangePasswordAtLogon = $false # Keine Passwortänderung erforderlich
            }

            # Optional: Füge zusätzliche Attribute hinzu, wenn sie angegeben sind
            if ($GivenName) { $userParams.GivenName = $GivenName }
            if ($Surname) { $userParams.Surname = $Surname }
            if ($Department) { $userParams.Department = $Department }
            if ($Title) { $userParams.Title = $Title }
            if ($Manager) { $userParams.Manager = $Manager }
            if ($Office) { $userParams.Office = $Office }
            if ($OfficePhone) { $userParams.OfficePhone = $OfficePhone }
            if ($Company) { $userParams.Company = $Company }
            if ($Description) { $userParams.Description = $Description }

            New-ADUser @userParams

            foreach ($group in $Groups) {
                Add-ADGroupMember -Identity $group -Members $SamAccountName
                Write-Verbose "Gruppe '$group' dem Benutzer '$SamAccountName' zugewiesen"
            }

            Write-CustomLog "Benutzer '$Name' erfolgreich erstellt in OU: $OU" -Level "INFO"
        }
    } catch {
        Write-CustomLog "Fehler beim Erstellen von Benutzer '$Name': $_" -Level "FEHLER"
    }
}

# Hauptausführungsblock
try {
    Write-CustomLog "Skript-Ausführung gestartet" -Level "INFO"

    # Interaktive Parameterabfrage, falls erforderlich
    if (-not $CsvPath) {
        $CsvPath = Read-Host "Bitte CSV-Dateipfad eingeben"
    }

    if ($CsvPath) {
        Import-ADUsersFromCSV -CsvPath $CsvPath
    }

    Write-CustomLog "Skript-Ausführung erfolgreich beendet" -Level "INFO"
} catch {
    Write-CustomLog "Unerwarteter Fehler bei der Skript-Ausführung: $_" -Level "FEHLER"
}
