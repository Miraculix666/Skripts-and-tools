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
        [string]$NewUserPassword
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
    
    $newUser = New-ADUser @newUserParams -PassThru
    Copy-ADGroupMembership -SourceUser $TemplateUser -TargetUser $newUser.SamAccountName
    
    return $newUser
}

# Hauptskript
#Vorlage für die CSV
# "TemplateUser","NewUserName","NewUserPassword","FirstName","LastName","Department","Title","Email"
# "john.doe","max.mustermann","Willkommen2025!","Max","Mustermann","IT","Junior Entwickler",""
# "jane.smith","anna.schmidt","Neustart2025!","Anna","Schmidt","Marketing","Marketing Spezialist",""


$inputMethod = Read-Host "Wählen Sie die Eingabemethode: 1) Interaktiv, 2) Vordefiniert, 3) CSV"

switch ($inputMethod) {
    "1" {
        $templateUser = Read-Host "Geben Sie den Benutzernamen der Vorlage ein"
        $newUserName = Read-Host "Geben Sie den Benutzernamen des neuen Benutzers ein"
        $newUserPassword = Read-Host "Geben Sie das Passwort für den neuen Benutzer ein" -AsSecureString
        $newUser = New-ADUserFromTemplate -TemplateUser $templateUser -NewUserName $newUserName -NewUserPassword ([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($newUserPassword)))
    }
    "2" {
        $templateUser = "vorlage_benutzer"
        $newUserName = "neuer_benutzer"
        $newUserPassword = "Sicheres_Passwort123!"
        $newUser = New-ADUserFromTemplate -TemplateUser $templateUser -NewUserName $newUserName -NewUserPassword $newUserPassword
    }
    "3" {
        $csvPath = Read-Host "Geben Sie den Pfad zur CSV-Datei ein"
        $users = Import-Csv $csvPath
        foreach ($user in $users) {
            $newUser = New-ADUserFromTemplate -TemplateUser $user.TemplateUser -NewUserName $user.NewUserName -NewUserPassword $user.NewUserPassword
        }
    }
}

Write-Host "Neue(r) Benutzer wurde(n) erstellt und Berechtigungen wurden kopiert."
