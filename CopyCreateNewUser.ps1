<#
.SYNOPSIS
  Erstellt neue AD-Benutzer durch Kopieren eines Template-Benutzers.
.DESCRIPTION
  Dieses Skript kann im interaktiven Modus einen einzelnen Benutzer oder im Batch-Modus alle neuen Benutzer aus einer CSV-Datei erstellen.
  Es kopiert dabei alle „copyable“ Attribute (z. B. Adresse, Abteilung, Titel, Manager) aus einem vorhandenen Template-Benutzer und überschreibt zwingende Attribute (wie Name, SamAccountName, UPN, Passwort, OU).
  Optional können Gruppenmitgliedschaften hinzugefügt werden. Fehlende mandatory Felder (z. B. bei CSV) werden erkannt und führen zu einer Logmeldung bzw. – im interaktiven Modus – zu einer Abfrage.
.NOTES
  Voraussetzung: Das ActiveDirectory-Modul muss geladen sein.
  Für den CSV-Import sollten die Spaltennamen u.a. enthalten: Name, SamAccountName, GivenName, Surname, UserPrincipalName, Password, Path – sowie optional z. B. StreetAddress, City, PostalCode, Department, Title, Office, Manager, Groups.
.INPUTS
  -TemplateUser  : (Optional) Der Identity-Name des Template-Benutzers.
  -CSVFile       : (Optional) Pfad zur CSV-Datei für Batchimport.
  -Interactive   : (Optional) Interaktiver Modus zur Erstellung eines einzelnen Benutzers.
.EXAMPLE
  .\New-ADUserFromTemplate.ps1 -TemplateUser "ctemplate" -Interactive
  Erstellt einen neuen Benutzer in interaktivem Modus unter Verwendung der Eigenschaften des Template-Benutzers "ctemplate".
  
  .\New-ADUserFromTemplate.ps1 -TemplateUser "ctemplate" -CSVFile "C:\it\newusers.csv"
  Liest aus der CSV-Datei und erstellt mehrere neue Benutzer.
#>

[CmdletBinding()]
param (
    [Parameter(Position=0)]
    [string]$TemplateUser,
    
    [Parameter(Position=1)]
    [string]$CSVFile,
    
    [Parameter()]
    [switch]$Interactive
)

# Logdatei definieren (im Skriptverzeichnis)
$LogFile = Join-Path -Path $PSScriptRoot -ChildPath "ADUserCreation.log"

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -FilePath $LogFile -Append
}

function Get-MandatoryValue {
    param(
        [string]$Prompt
    )
    do {
        $value = Read-Host $Prompt
        if ([string]::IsNullOrWhiteSpace($value)) {
            Write-Host "Dieser Wert wird benötigt. Bitte geben Sie einen Wert ein." -ForegroundColor Yellow
        }
    } until (-not [string]::IsNullOrWhiteSpace($value))
    return $value
}

function CreateUserFromTemplate {
    param (
        [PSCustomObject]$UserData,    # Neue Benutzerangaben (mandatory und optionale Felder)
        [object]$TemplateObj          # Template-Benutzer (kann auch $null sein)
    )

    # Definieren Sie die optional kopierbaren Attribute. Falls im CSV/interaktiven Input kein neuer Wert angegeben wurde, wird der Wert aus dem Template übernommen.
    $optionalAttrs = @{}
    foreach ($attr in "StreetAddress", "City", "PostalCode", "Department", "Title", "Office", "Manager") {
        if (-not [string]::IsNullOrWhiteSpace($UserData.$attr)) {
            $optionalAttrs[$attr] = $UserData.$attr
        }
        elseif ($TemplateObj -and ($TemplateObj.$attr)) {
            $optionalAttrs[$attr] = $TemplateObj.$attr
        }
    }

    # Mandatory Attribute – zwingend neu zu setzen (Name, SamAccountName, GivenName, Surname, UserPrincipalName, Passwort, und OU (Path))
    $mandatoryAttrs = @{
        Name               = $UserData.Name
        SamAccountName     = $UserData.SamAccountName
        GivenName          = $UserData.GivenName
        Surname            = $UserData.Surname
        UserPrincipalName  = $UserData.UserPrincipalName
        AccountPassword    = $UserData.AccountPassword  # Bereits als SecureString
        Path               = $UserData.Path
        Enabled            = $true
        ChangePasswordAtLogon = $true
    }

    # Zusammenführen beider Hashtables
    $newUserParams = $mandatoryAttrs + $optionalAttrs

    try {
        if ($TemplateObj) {
            Write-Log "Erstelle Benutzer $($newUserParams.Name) mittels Template-Kopie..."
            # Mit -Instance werden kopierbare Eigenschaften aus dem Template übernommen. Überschriebene zwingende Attribute werden aus $newUserParams übernommen.
            New-ADUser -Instance $TemplateObj @newUserParams
        }
        else {
            Write-Log "Erstelle Benutzer $($newUserParams.Name) ohne Template..."
            New-ADUser @newUserParams
        }
        Write-Log "Benutzer $($newUserParams.Name) wurde erfolgreich angelegt."
    }
    catch {
        Write-Log "Fehler bei der Erstellung von Benutzer $($newUserParams.Name): $_"
        Write-Host "FEHLER: $_" -ForegroundColor Red
    }

    # Hinzufügen von Gruppenmitgliedschaften (wenn in UserData.Groups als kommaseparierte Liste angegeben)
    if (-not [string]::IsNullOrWhiteSpace($UserData.Groups)) {
        $groups = $UserData.Groups -split ","
        foreach ($grp in $groups) {
            $grp = $grp.Trim()
            if ($grp) {
                try {
                    Add-ADGroupMember -Identity $grp -Members $newUserParams.SamAccountName
                    Write-Log "Benutzer $($newUserParams.Name) wurde zur Gruppe '$grp' hinzugefügt."
                }
                catch {
                    Write-Log "Fehler beim Hinzufügen von Benutzer $($newUserParams.Name) zur Gruppe '$grp': $_"
                }
            }
        }
    }
}

# Modul laden
Import-Module ActiveDirectory

# Template-Benutzer laden, falls angegeben
if ($TemplateUser) {
    try {
        $TemplateObj = Get-ADUser -Identity $TemplateUser -Properties *
        Write-Log "Template-Benutzer '$TemplateUser' wurde erfolgreich geladen."
    }
    catch {
        Write-Host "Template-Benutzer '$TemplateUser' nicht gefunden. Fortsetzung ohne Template-Kopie." -ForegroundColor Yellow
        Write-Log "Template-Benutzer '$TemplateUser' nicht gefunden: $_. Fortsetzung ohne Template."
        $TemplateObj = $null
    }
}
else {
    $TemplateObj = $null
}

# Auswahl zwischen CSV-gestütztem Batch-Modus und interaktivem Modus
if ($CSVFile) {
    Write-Log "Starte Batch-Modus mit CSV-Datei: $CSVFile"
    try {
        $users = Import-Csv -Path $CSVFile
        foreach ($user in $users) {
            # Prüfe mandatory Felder: Name, SamAccountName, GivenName, Surname, UserPrincipalName, Password, Path
            $fehlendeFelder = @()
            if (-not $user.Name)              { $fehlendeFelder += "Name" }
            if (-not $user.SamAccountName)      { $fehlendeFelder += "SamAccountName" }
            if (-not $user.GivenName)           { $fehlendeFelder += "GivenName" }
            if (-not $user.Surname)             { $fehlendeFelder += "Surname" }
            if (-not $user.UserPrincipalName)   { $fehlendeFelder += "UserPrincipalName" }
            if (-not $user.Password)            { $fehlendeFelder += "Password" }
            if (-not $user.Path)                { $fehlendeFelder += "OU (Path)" }

            if ($fehlendeFelder.Count -gt 0) {
                $msg = "Für Benutzer '$($user.Name)' fehlen mandatory Felder: $($fehlendeFelder -join ', '). – Dieser Datensatz wird übersprungen."
                Write-Log $msg
                Write-Host $msg -ForegroundColor Yellow
                continue
            }
            
            # Konvertiere das Passwort in einen SecureString
            $securePass = ConvertTo-SecureString $user.Password -AsPlainText -Force
            $user | Add-Member -MemberType NoteProperty -Name AccountPassword -Value $securePass
            
            CreateUserFromTemplate -UserData $user -TemplateObj $TemplateObj
        }
    }
    catch {
        Write-Log "Fehler beim laden der CSV-Datei: $_"
        Write-Host "FEHLER beim CSV-Import: $_" -ForegroundColor Red
    }
}
elseif ($Interactive) {
    Write-Host "Interaktiver Modus: Neuer Benutzer wird erstellt." -ForegroundColor Cyan
    # Erfassen der zwingenden Angaben
    $userData = [PSCustomObject]@{
        Name              = Get-MandatoryValue "Bitte vollständigen Namen eingeben (Name):"
        SamAccountName    = Get-MandatoryValue "Bitte SamAccountName eingeben:"
        GivenName         = Get-MandatoryValue "Bitte Vornamen eingeben (GivenName):"
        Surname           = Get-MandatoryValue "Bitte Nachnamen eingeben (Surname):"
        UserPrincipalName = Get-MandatoryValue "Bitte User Principal Name eingeben (z.B. user@domain.com):"
        Password          = Get-MandatoryValue "Bitte initiales Passwort eingeben:"
        Path              = Get-MandatoryValue "Bitte OU-Pfad angeben (z.B. OU=Users,DC=domain,DC=com):"
        StreetAddress     = Read-Host "Straßenadresse (optional):"
        City              = Read-Host "Stadt (optional):"
        PostalCode        = Read-Host "Postleitzahl (optional):"
        Department        = Read-Host "Abteilung (optional):"
        Title             = Read-Host "Jobtitel (optional):"
        Office            = Read-Host "Office (optional):"
        Manager           = Read-Host "Manager (optional, DN):"
        Groups            = Read-Host "Gruppenmitgliedschaften (optional, kommasepariert):"
    }
    # Passwort konvertieren
    $userData | Add-Member -MemberType NoteProperty -Name AccountPassword -Value (ConvertTo-SecureString $userData.Password -AsPlainText -Force)
    
    CreateUserFromTemplate -UserData $userData -TemplateObj $TemplateObj
}
else {
    Write-Host "Bitte entweder den interaktiven Modus (-Interactive) nutzen oder den Pfad zu einer CSV-Datei (--CSVFile) angeben." -ForegroundColor Red
    Write-Log "Kein Ausführungsmodus angegeben. Skript wird beendet."
}
