<#
.SYNOPSIS
    Erstellt AD-Benutzer basierend auf einem Vorlagenbenutzer und CSV-Daten.

.DESCRIPTION
    Dieses Skript erstellt neue AD-Benutzer basierend auf einem Vorlagenbenutzer und CSV-Daten.
    Es unterstützt auch das Auslesen der Daten des Vorlagenbenutzers in eine CSV-Datei.

.PARAMETER CsvPath
    Pfad zur CSV-Datei mit den Benutzerdaten.

.PARAMETER TemplateUser
    SAMAccountName des Vorlagenbenutzers.

.PARAMETER ExportTemplateOnly
    Schalter zum Auslesen der Daten des Vorlagenbenutzers in eine CSV-Datei.

.EXAMPLE
    .\CopyCreateNewUser.ps1 -CsvPath "C:\Users.csv" -TemplateUser "TemplateUser"

.EXAMPLE
    .\CopyCreateNewUser.ps1 -TemplateUser "TemplateUser" -ExportTemplateOnly
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [string]$CsvPath,

    [Parameter(Mandatory=$true)]
    [string]$TemplateUser,

    [switch]$ExportTemplateOnly
)

# Importiere das Active Directory-Modul
Import-Module ActiveDirectory

# Funktion zum Erstellen des Log-Verzeichnisses
function Create-LogDirectory {
    $logDir = "C:\ADUserCreationLogs"
    if (!(Test-Path $logDir)) {
        New-Item -ItemType Directory -Force -Path $logDir
    }
    return $logDir
}

# Funktion zum Schreiben von Log-Einträgen
function Write-Log {
    param (
        [string]$Message
    )
    $logDir = Create-LogDirectory
    $logFile = Join-Path $logDir "ADUserCreation_$(Get-Date -Format 'yyyyMMdd').log"
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -FilePath $logFile -Append
    Write-Verbose $Message
}

# Funktion zum Exportieren der Vorlagenbenutzerdaten
function Export-TemplateUserData {
    param (
        [string]$TemplateUser,
        [string]$CsvPath
    )
    $user = Get-ADUser -Identity $TemplateUser -Properties *
    $userProperties = $user | Select-Object SamAccountName, GivenName, Surname, Name, DisplayName, Description, Office, Department, Title, Company, EmailAddress, StreetAddress, City, State, PostalCode, Country, OfficePhone
    $userProperties | Export-Csv -Path $CsvPath -NoTypeInformation -Delimiter ";"
    Write-Log "Vorlagenbenutzerdaten wurden in $CsvPath exportiert."
}

# Funktion zum Erstellen eines neuen AD-Benutzers
function New-ADUserFromTemplate {
    param (
        [hashtable]$UserData,
        [string]$TemplateUser
    )
    $template = Get-ADUser -Identity $TemplateUser -Properties *
    $securePassword = ConvertTo-SecureString $UserData.Password -AsPlainText -Force

    $newUserParams = @{
        SamAccountName = $UserData.SamAccountName
        UserPrincipalName = "$($UserData.SamAccountName)@$((Get-ADDomain).DNSRoot)"
        Name = $UserData.Name
        GivenName = $UserData.GivenName
        Surname = $UserData.Surname
        DisplayName = $UserData.DisplayName
        Description = $UserData.Description
        Office = $UserData.Office
        Department = $UserData.Department
        Title = $UserData.Title
        Company = $UserData.Company
        EmailAddress = $UserData.EmailAddress
        StreetAddress = $UserData.StreetAddress
        City = $UserData.City
        State = $UserData.State
        PostalCode = $UserData.PostalCode
        Country = $UserData.Country
        OfficePhone = $UserData.OfficePhone
        AccountPassword = $securePassword
        Enabled = $true
        Path = ($template.DistinguishedName -split ',', 2)[1]
        Instance = $template
    }

    try {
        New-ADUser @newUserParams
        Write-Log "Benutzer $($UserData.SamAccountName) wurde erfolgreich erstellt."
    }
    catch {
        Write-Log "Fehler beim Erstellen des Benutzers $($UserData.SamAccountName): $_"
    }
}

# Hauptlogik
if ($ExportTemplateOnly) {
    Export-TemplateUserData -TemplateUser $TemplateUser -CsvPath $CsvPath
}
else {
    $users = Import-Csv -Path $CsvPath -Delimiter ";"
    foreach ($user in $users) {
        if ([string]::IsNullOrEmpty($user.Password)) {
            Write-Log "Fehler: Passwort für Benutzer $($user.SamAccountName) fehlt. Dieser Benutzer wird übersprungen."
            continue
        }
        New-ADUserFromTemplate -UserData $user -TemplateUser $TemplateUser
    }
}

Write-Log "Skript-Ausführung abgeschlossen."
