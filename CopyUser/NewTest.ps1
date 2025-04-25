[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [string]$CSVPath,
    
    [Parameter(Mandatory=$false)]
    [string]$TemplateUser,
    
    [Parameter(Mandatory=$false)]
    [switch]$ExportTemplateOnly,
    
    [Parameter(Mandatory=$false)]
    [switch]$NoConfirmation,
    
    [Parameter(Mandatory=$false)]
    [string]$TargetOU,
    
    [Parameter(Mandatory=$false)]
    [switch]$Interactive
)

# Funktionen für Logging und Ausgabe
function Write-LogMessage {
    param($Message, $Type = "Info")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Type] $Message"
    
    switch ($Type) {
        "Error"   { Write-Host $logMessage -ForegroundColor Red }
        "Warning" { Write-Host $logMessage -ForegroundColor Yellow }
        "Success" { Write-Host $logMessage -ForegroundColor Green }
        default   { Write-Host $logMessage -ForegroundColor Cyan }
    }
    
    Add-Content -Path ".\ADUser_Creation_Log.txt" -Value $logMessage
}

# Funktion zum Exportieren der Template-Daten
function Export-TemplateUser {
    param([string]$Username)
    
    try {
        $user = Get-ADUser -Identity $Username -Properties *
        $properties = @{
            "SamAccountName" = ""
            "GivenName" = $user.GivenName
            "Surname" = $user.Surname
            "Password" = ""
            "Department" = $user.Department
            "Title" = $user.Title
            "Company" = $user.Company
            "Office" = $user.Office
            "StreetAddress" = $user.StreetAddress
            "City" = $user.City
            "PostalCode" = $user.PostalCode
            "Country" = $user.Country
            "TelephoneNumber" = $user.TelephoneNumber
        }
        
        $exportPath = ".\TemplateUser_Export.csv"
        [PSCustomObject]$properties | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
        Write-LogMessage "Template-Benutzer wurde exportiert nach: $exportPath" -Type "Success"
    }
    catch {
        Write-LogMessage "Fehler beim Exportieren des Template-Benutzers: $_" -Type "Error"
        exit
    }
}

# Funktion zur Benutzervalidierung
function Confirm-UserCreation {
    param($UserData)
    
    Write-Host "`n=== Neue Benutzer-Details ===" -ForegroundColor Cyan
    $UserData.PSObject.Properties | ForEach-Object {
        Write-Host "$($_.Name): $($_.Value)" -ForegroundColor Yellow
    }
    
    if (-not $NoConfirmation) {
        $confirm = Read-Host "`nBenutzer erstellen? (J/N)"
        return $confirm -eq "J"
    }
    return $true
}

# Hauptfunktion zur Benutzererstellung
function New-CustomADUser {
    param($UserData)
    
    try {
        $securePassword = ConvertTo-SecureString $UserData.Password -AsPlainText -Force
        
        $params = @{
            SamAccountName = $UserData.SamAccountName
            UserPrincipalName = "$($UserData.SamAccountName)@$((Get-ADDomain).DNSRoot)"
            GivenName = $UserData.GivenName
            Surname = $UserData.Surname
            Name = "$($UserData.GivenName) $($UserData.Surname)"
            DisplayName = "$($UserData.GivenName) $($UserData.Surname)"
            AccountPassword = $securePassword
            Enabled = $true
            Path = $TargetOU
        }
        
        # Optionale Parameter hinzufügen
        @("Department", "Title", "Company", "Office", "StreetAddress", "City", "PostalCode", "Country", "TelephoneNumber") | ForEach-Object {
            if ($UserData.$_) {
                $params[$_] = $UserData.$_
            }
        }
        
        New-ADUser @params
        Write-LogMessage "Benutzer $($UserData.SamAccountName) wurde erfolgreich erstellt" -Type "Success"
    }
    catch {
        Write-LogMessage "Fehler bei der Erstellung von $($UserData.SamAccountName): $_" -Type "Error"
    }
}

# Hauptprogramm
try {
    # Modul-Check
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-LogMessage "Active Directory PowerShell-Modul nicht gefunden!" -Type "Error"
        exit
    }
    
    # OU-Validierung
    if (-not $TargetOU) {
        $TargetOU = Read-Host "Bitte geben Sie die Ziel-OU an (z.B. 'OU=Users,DC=domain,DC=com')"
    }
    if (-not (Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$TargetOU'")) {
        Write-LogMessage "Die angegebene OU existiert nicht!" -Type "Error"
        exit
    }
    
    # Template-Export
    if ($TemplateUser -and $ExportTemplateOnly) {
        Export-TemplateUser -Username $TemplateUser
        exit
    }
    
    # Benutzerverarbeitung
    if ($CSVPath) {
        $users = Import-Csv -Path $CSVPath -Encoding UTF8
    }
    elseif ($Interactive) {
        # Interaktive Benutzereingabe
        $userData = @{
            SamAccountName = Read-Host "SamAccountName"
            GivenName = Read-Host "Vorname"
            Surname = Read-Host "Nachname"
            Password = Read-Host "Passwort"
        }
        $users = @([PSCustomObject]$userData)
    }
    
    foreach ($user in $users) {
        if (-not $user.Password) {
            Write-LogMessage "Kein Passwort für Benutzer $($user.SamAccountName) angegeben" -Type "Error"
            continue
        }
        
        if (Confirm-UserCreation -UserData $user) {
            New-CustomADUser -UserData $user
        }
    }
}
catch {
    Write-LogMessage "Unerwarteter Fehler: $_" -Type "Error"
}
