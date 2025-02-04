<#
.SYNOPSIS
Erstellt neue AD-Benutzer basierend auf vorhandenen Vorlagen und überprüft optional die Erstellung.

.DESCRIPTION
Dieses Skript erstellt neue AD-Benutzer, indem es vorhandene Benutzer als Vorlagen verwendet und deren Attribute und Berechtigungen kopiert. Optional kann es die erfolgreiche Erstellung und korrekte Verrechtung überprüfen.

.PARAMETER CSV
Der Pfad zur CSV-Datei mit den Benutzerdaten.

.PARAMETER TemplateUser
Der Benutzername des Vorlagenbenutzers.

.PARAMETER NewUserName
Der Benutzername für den neuen AD-Benutzer.

.PARAMETER NewUserPassword
Das Passwort für den neuen AD-Benutzer.

.PARAMETER Verify
Schalter zur Aktivierung der Überprüfung nach der Benutzererstellung.

.EXAMPLE
.\New-ADUserFromTemplate.ps1
.\New-ADUserFromTemplate.ps1 -CSV "C:\Pfad\zu\ad_user_creation_template.csv"
.\New-ADUserFromTemplate.ps1 -TemplateUser "john.doe" -NewUserName "jane.smith" -NewUserPassword "P@ssw0rd123!" -Verify

.NOTES
Beispiel-CSV-Inhalt:
TemplateUser,NewUserName,NewUserPassword,FirstName,LastName,Department,Title,Email
john.doe,max.mustermann,Willkommen2025!,Max,Mustermann,IT,Junior Entwickler,max.mustermann@unternehmen.de
jane.smith,anna.schmidt,Neustart2025!,Anna,Schmidt,Marketing,Marketing Spezialist,anna.schmidt@unternehmen.de
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [string]$CSV,

    [Parameter(Mandatory=$false)]
    [string]$TemplateUser,

    [Parameter(Mandatory=$false)]
    [string]$NewUserName,

    [Parameter(Mandatory=$false)]
    [string]$NewUserPassword,

    [Parameter(Mandatory=$false)]
    [switch]$Verify
)

# Importiere das Active Directory-Modul
Import-Module ActiveDirectory

# Funktion zum Kopieren der Gruppenmitgliedschaften
function Copy-ADGroupMembership {
    param (
        [string]$SourceUser,
        [string]$TargetUser
    )
    
    $groups = Get-ADUser -Identity $SourceUser -Properties MemberOf | Select-Object -ExpandProperty MemberOf
    foreach ($group in $groups) {
        Add-ADGroupMember -Identity $group -Members $TargetUser
    }
}

# Funktion zum Erstellen eines neuen Benutzers basierend auf einer Vorlage
function New-ADUserFromTemplate {
    param (
        [string]$TemplateUser,
        [string]$NewUserName,
        [string]$NewUserPassword,
        [hashtable]$AdditionalProperties
    )
    
    $template = Get-ADUser -Identity $TemplateUser -Properties *
    $newUserParams = @{
        Instance = $template
        Name = $NewUserName
        SamAccountName = $NewUserName
        UserPrincipalName = "$NewUserName@$((Get-ADDomain).DNSRoot)"
        Enabled = $true
        AccountPassword = (ConvertTo-SecureString -AsPlainText $NewUserPassword -Force)
    }
    
    # Füge zusätzliche Eigenschaften hinzu, falls vorhanden
    foreach ($key in $AdditionalProperties.Keys) {
        $newUserParams[$key] = $AdditionalProperties[$key]
    }
    
    try {
        $newUser = New-ADUser @newUserParams -PassThru
        Copy-ADGroupMembership -SourceUser $TemplateUser -TargetUser $newUser.SamAccountName
        Write-Host "Benutzer $NewUserName wurde erfolgreich erstellt." -ForegroundColor Green
        return $newUser
    }
    catch {
        Write-Host "Fehler beim Erstellen des Benutzers $NewUserName: $_" -ForegroundColor Red
        return $null
    }
}

# Funktion zur Überprüfung der Benutzererstellung und Verrechtung
function Verify-ADUser {
    param (
        [string]$TemplateUser,
        [string]$NewUserName
    )

    $template = Get-ADUser -Identity $TemplateUser -Properties *
    $newUser = Get-ADUser -Identity $NewUserName -Properties *

    Write-Host "Überprüfe Benutzer $NewUserName..."

    # Überprüfe, ob der Benutzer existiert
    if ($null -eq $newUser) {
        Write-Host "Fehler: Benutzer $NewUserName wurde nicht erstellt." -ForegroundColor Red
        return
    }

    Write-Host "Benutzer $NewUserName existiert." -ForegroundColor Green

    # Überprüfe Gruppenmitgliedschaften
    $templateGroups = Get-ADUser -Identity $TemplateUser -Properties MemberOf | Select-Object -ExpandProperty MemberOf
    $newUserGroups = Get-ADUser -Identity $NewUserName -Properties MemberOf | Select-Object -ExpandProperty MemberOf

    $missingGroups = Compare-Object -ReferenceObject $templateGroups -DifferenceObject $newUserGroups | Where-Object { $_.SideIndicator -eq '<=' } | Select-Object -ExpandProperty InputObject

    if ($missingGroups) {
        Write-Host "Warnung: Folgende Gruppen fehlen beim neuen Benutzer:" -ForegroundColor Yellow
        $missingGroups | ForEach-Object { Write-Host "- $($_ -split ',')[0]" -ForegroundColor Yellow }
    } else {
        Write-Host "Gruppenmitgliedschaften sind korrekt." -ForegroundColor Green
    }

    # Überprüfe andere relevante Attribute
    $relevantAttributes = @('Department', 'Title', 'Email', 'GivenName', 'Surname', 'DisplayName')
    foreach ($attr in $relevantAttributes) {
        if ($newUser.$attr -ne $template.$attr) {
            Write-Host "Warnung: $attr stimmt nicht überein. Vorlage: $($template.$attr), Neu: $($newUser.$attr)" -ForegroundColor Yellow
        } else {
            Write-Host "$attr ist korrekt." -ForegroundColor Green
        }
    }
}

# Hauptskript
$createdUsers = @()

if ($CSV) {
    # CSV-Modus
    $users = Import-Csv $CSV
    foreach ($user in $users) {
        $additionalProps = @{}
        foreach ($prop in $user.PSObject.Properties) {
            if ($prop.Name -notin @('TemplateUser', 'NewUserName', 'NewUserPassword')) {
                $additionalProps[$prop.Name] = $prop.Value
            }
        }
        $newUser = New-ADUserFromTemplate -TemplateUser $user.TemplateUser -NewUserName $user.NewUserName -NewUserPassword $user.NewUserPassword -AdditionalProperties $additionalProps
        if ($newUser) {
            $createdUsers += @{TemplateUser = $user.TemplateUser; NewUserName = $user.NewUserName}
        }
    }
}
elseif ($TemplateUser -or $NewUserName -or $NewUserPassword) {
    # Überprüfung und Abfrage fehlender Parameter
    if (-not $TemplateUser) {
        $TemplateUser = Read-Host "Geben Sie den Benutzernamen der Vorlage ein"
    }
    if (-not $NewUserName) {
        $NewUserName = Read-Host "Geben Sie den Benutzernamen des neuen Benutzers ein"
    }
    if (-not $NewUserPassword) {
        $NewUserPassword = Read-Host "Geben Sie das Passwort für den neuen Benutzer ein" -AsSecureString
        $NewUserPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($NewUserPassword))
    }
    
    $additionalProps = @{}
    $newUser = New-ADUserFromTemplate -TemplateUser $TemplateUser -NewUserName $NewUserName -NewUserPassword $NewUserPassword -AdditionalProperties $additionalProps
    if ($newUser) {
        $createdUsers += @{TemplateUser = $TemplateUser; NewUserName = $NewUserName}
    }
}
else {
    # Interaktiver Modus
    $TemplateUser = Read-Host "Geben Sie den Benutzernamen der Vorlage ein"
    $NewUserName = Read-Host "Geben Sie den Benutzernamen des neuen Benutzers ein"
    $NewUserPassword = Read-Host "Geben Sie das Passwort für den neuen Benutzer ein" -AsSecureString
    $NewUserPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($NewUserPassword))
    
    $additionalProps = @{}
    $newUser = New-ADUserFromTemplate -TemplateUser $TemplateUser -NewUserName $NewUserName -NewUserPassword $NewUserPassword -AdditionalProperties $additionalProps
    if ($newUser) {
        $createdUsers += @{TemplateUser = $TemplateUser; NewUserName = $NewUserName}
    }
}

Write-Host "Erstellung der neuen Benutzer abgeschlossen."

if ($Verify -or ($createdUsers.Count -gt 0 -and (Read-Host "Möchten Sie die erstellten Benutzer überprüfen? (J/N)") -eq 'J')) {
    foreach ($user in $createdUsers) {
        Verify-ADUser -TemplateUser $user.TemplateUser -NewUserName $user.NewUserName
    }
}
