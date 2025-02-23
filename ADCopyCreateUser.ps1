# Active Directory Benutzerverwaltungsskript
# Autor: [Ihr Name]
# Beschreibung:
# - Exportiert bestehende AD-Benutzer in eine CSV-Datei
# - Erstellt neue Benutzer entweder interaktiv, mit Parametern oder per CSV
# - Setzt Benutzergruppen basierend auf Vorlage

Import-Module ActiveDirectory

function Export-ADUsers {
    param (
        [string]$ExportPath = "ADUsersExport.csv"
    )
    
    Write-Host "Exportiere Benutzer aus dem Active Directory..." -ForegroundColor Cyan
    
    try {
        Get-ADUser -Filter * -Properties SamAccountName, UserPrincipalName, Name, MemberOf |
        Select-Object SamAccountName, UserPrincipalName, Name, @{Name='Groups'; Expression={$_.MemberOf -join ';'}} |
        Export-Csv -Path $ExportPath -NoTypeInformation

        Write-Host "Benutzerdaten erfolgreich exportiert: $ExportPath" -ForegroundColor Green
    } catch {
        Write-Host "Fehler beim Export: $_" -ForegroundColor Red
    }
}

function Create-ADUser {
    param (
        [string]$SamAccountName,
        [string]$UserPrincipalName,
        [string]$Name,
        [string]$OU,
        [string[]]$Groups,
        [string]$Password = "Passwort123!"
    )

    Write-Host "Erstelle Benutzer: $Name ($SamAccountName)..." -ForegroundColor Cyan
    
    try {
        New-ADUser -Name $Name `
                   -SamAccountName $SamAccountName `
                   -UserPrincipalName $UserPrincipalName `
                   -Path $OU `
                   -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -Force) `
                   -Enabled $true

        foreach ($group in $Groups) {
            Add-ADGroupMember -Identity $group -Members $SamAccountName
        }

        Write-Host "Benutzer $Name wurde erfolgreich erstellt." -ForegroundColor Green
    } catch {
        Write-Host "Fehler beim Erstellen von $Name: $_" -ForegroundColor Red
    }
}

function Create-ADUsersFromCSV {
    param (
        [string]$CsvPath
    )
    
    Write-Host "Erstelle Benutzer aus CSV: $CsvPath..." -ForegroundColor Cyan
    
    if (-Not (Test-Path $CsvPath)) {
        Write-Host "Fehler: Datei $CsvPath nicht gefunden!" -ForegroundColor Red
        return
    }
    
    $users = Import-Csv -Path $CsvPath
    foreach ($user in $users) {
        Create-ADUser -SamAccountName $user.SamAccountName `
                      -UserPrincipalName $user.UserPrincipalName `
                      -Name $user.Name `
                      -OU $user.OU `
                      -Groups ($user.Groups -split ';') `
                      -Password $user.Password
    }
}

function Show-Menu {
    Write-Host "==================================" -ForegroundColor Yellow
    Write-Host " Active Directory Benutzerverwaltung " -ForegroundColor Yellow
    Write-Host "==================================" -ForegroundColor Yellow
    Write-Host "[1] Exportiere AD-Benutzer nach CSV"
    Write-Host "[2] Erstelle Benutzer interaktiv"
    Write-Host "[3] Erstelle Benutzer per CSV"
    Write-Host "[0] Beenden"
}

# Hauptprogramm
while ($true) {
    Show-Menu
    $choice = Read-Host "Bitte eine Option wählen"
    switch ($choice) {
        "1" { Export-ADUsers }
        "2" {
            $SamAccountName = Read-Host "SamAccountName"
            $UserPrincipalName = Read-Host "UserPrincipalName"
            $Name = Read-Host "Vollständiger Name"
            $OU = Read-Host "Organizational Unit (OU)"
            $Groups = Read-Host "Gruppen (mit ; trennen)" -split ';'
            $Password = Read-Host "Passwort (leer für Standard)" -AsSecureString
            Create-ADUser -SamAccountName $SamAccountName -UserPrincipalName $UserPrincipalName -Name $Name -OU $OU -Groups $Groups -Password $Password
        }
        "3" {
            $CsvPath = Read-Host "Pfad zur CSV-Datei"
            Create-ADUsersFromCSV -CsvPath $CsvPath
        }
        "0" { break }
        default { Write-Host "Ungültige Auswahl!" -ForegroundColor Red }
    }
}
