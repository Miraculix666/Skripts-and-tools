# Benutzername	E-Mail	Vorname	Nachname	Abteilung	Position	Vorgesetzter	Büro	Telefon	Firma	Beschreibung	OU	Gruppen
# Funktion zum Importieren von Benutzern aus CSV
function Import-ADUsersFromCSV {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()

    if (-not (Test-Path $CsvPath)) {
        Write-CustomLog "CSV-Datei nicht gefunden: $CsvPath" -Level "FEHLER"
        return
    }

    Write-CustomLog "Importiere Benutzer aus CSV: $CsvPath" -Level "INFO"

    try {
        $templateOUs = Get-TemplateUserOUs -TemplateUser $TemplateUser
        $users = Import-Csv -Path $CsvPath -Delimiter ";" -Encoding UTF8

        foreach ($user in $users) {
            # Bestimme OU basierend auf Priorität: CSV > DefaultOU > Zufällige Template OU
            $ou = if ($user.OU -and ($templateOUs -contains $user.OU)) {
                $user.OU
            } elseif ($DefaultOU) {
                $DefaultOU
            } else {
                $templateOUs | Get-Random
            }

            if ($PSCmdlet.ShouldProcess($user.Name, "Benutzer aus CSV erstellen")) {
                $params = @{
                    SamAccountName = $user.Benutzername
                    UserPrincipalName = $user.EMail
                    Name = ($user.Vorname + " " + $user.Nachname).Trim()
                    OU = $ou
                    Groups = ($user.Gruppen -split ';')
                }

                # Füge optionale Parameter hinzu, wenn sie im CSV existieren
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
