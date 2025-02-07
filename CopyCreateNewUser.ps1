[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$CsvPath,
    
    [Parameter(Mandatory=$false)]
    [string]$TemplateUser,
    
    [Parameter(Mandatory=$false)]
    [switch]$ExportTemplateOnly,
    
    [Parameter(Mandatory=$false)]
    [string]$TargetOU
)

# Funktionen für Logging und Ausgabe
function Write-LogMessage {
    param(
        [string]$Message,
        [ValidateSet('Info','Warning','Error','Success')]
        [string]$Type = 'Info'
    )
    
    $colors = @{
        'Info' = 'Cyan'
        'Warning' = 'Yellow'
        'Error' = 'Red'
        'Success' = 'Green'
    }
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] " -NoNewline
    Write-Host $Message -ForegroundColor $colors[$Type]
    
    # Logging in Datei
    $logPath = ".\ADUser_Creation_Log.txt"
    "[$timestamp] [$Type] $Message" | Out-File -FilePath $logPath -Append
}

# Funktion zum Exportieren der Template-Daten
function Export-TemplateUserData {
    param([string]$TemplateUser)
    
    try {
        $user = Get-ADUser -Identity $TemplateUser -Properties *
        $exportProperties = @(
            'GivenName','Surname','Department','Title',
            'City','Country','Company','Office'
        )
        
        $userData = $user | Select-Object $exportProperties
        $userData | Export-Csv -Path ".\TemplateUser_Export.csv" -NoTypeInformation -Encoding UTF8
        Write-LogMessage "Template-Daten wurden exportiert nach TemplateUser_Export.csv" -Type Success
    }
    catch {
        Write-LogMessage "Fehler beim Exportieren der Template-Daten: $_" -Type Error
        exit
    }
}

# Hauptfunktion zur Benutzerverarbeitung
function Process-UserCreation {
    param(
        [hashtable]$UserData,
        [string]$TargetOU
    )
    
    $securePassword = ConvertTo-SecureString $UserData.Password -AsPlainText -Force
    
    $newUserParams = @{
        Name = "$($UserData.GivenName) $($UserData.Surname)"
        GivenName = $UserData.GivenName
        Surname = $UserData.Surname
        SamAccountName = $UserData.SamAccountName
        UserPrincipalName = "$($UserData.SamAccountName)@$($env:USERDNSDOMAIN)"
        AccountPassword = $securePassword
        Enabled = $true
        Path = $TargetOU
    }
    
    # Optionale Parameter hinzufügen wenn vorhanden
    @('Department','Title','City','Country','Company','Office') | ForEach-Object {
        if ($UserData.$_) {
            $newUserParams[$_] = $UserData.$_
        }
    }
    
    try {
        New-ADUser @newUserParams
        Write-LogMessage "Benutzer $($UserData.SamAccountName) erfolgreich erstellt" -Type Success
    }
    catch {
        Write-LogMessage "Fehler beim Erstellen von $($UserData.SamAccountName): $_" -Type Error
    }
}

# Hauptprogramm
Write-LogMessage "Starte AD-Benutzerverarbeitung" -Type Info

# OU-Validierung
if (-not $TargetOU) {
    $TargetOU = Read-Host "Bitte geben Sie die Ziel-OU an (z.B. 'OU=Users,DC=domain,DC=com')"
}
if (-not (Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$TargetOU'")) {
    Write-LogMessage "Die angegebene OU existiert nicht: $TargetOU" -Type Error
    exit
}

# Template-Export wenn gewünscht
if ($ExportTemplateOnly -and $TemplateUser) {
    Export-TemplateUserData -TemplateUser $TemplateUser
    exit
}

# Verarbeitungsmodus bestimmen
if ($CsvPath) {
    # CSV-Modus
    $users = Import-Csv -Path $CsvPath -Encoding UTF8
    foreach ($user in $users) {
        if (-not $user.Password) {
            Write-LogMessage "Kein Passwort für Benutzer $($user.SamAccountName) angegeben" -Type Error
            continue
        }
        Process-UserCreation -UserData $user -TargetOU $TargetOU
    }
}
else {
    # Interaktiver Modus
    $userData = @{}
    $userData.GivenName = Read-Host "Vorname"
    $userData.Surname = Read-Host "Nachname"
    $userData.SamAccountName = Read-Host "SAM Account Name"
    $userData.Password = Read-Host "Passwort"
    
    if ($TemplateUser) {
        $template = Get-ADUser -Identity $TemplateUser -Properties *
        @('Department','Title','City','Country','Company','Office') | ForEach-Object {
            $userData[$_] = $template.$_
        }
    }
    
    Process-UserCreation -UserData $userData -TargetOU $TargetOU
}

Write-LogMessage "Verarbeitung abgeschlossen" -Type Success
