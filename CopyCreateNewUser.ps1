<#
.SYNOPSIS
    Erstellt AD-Benutzer als exakte Kopien von Vorlagenbenutzern.

.DESCRIPTION
    Dieses Skript erstellt neue AD-Benutzer als Klone vorhandener Vorlagen,
    inklusive aller Attribute und Gruppenmitgliedschaften.

.PARAMETER TemplateUser
    SAMAccountName des Vorlagenbenutzers

.PARAMETER NewUserName
    SAMAccountName für den neuen Benutzer

.PARAMETER NewUserPassword
    Initialpasswort für den neuen Benutzer

.PARAMETER Verify
    Aktiviert die Überprüfung nach der Erstellung

.EXAMPLE
    .\CopyCreateNewUser.ps1 -TemplateUser "L7111101" -NewUserName "L7111115" -NewUserPassword "P2f2aL5!01" -Verify
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory = $true)]
    [string]$TemplateUser,

    [Parameter(Mandatory = $true)]
    [string]$NewUserName,

    [Parameter(Mandatory = $true)]
    [string]$NewUserPassword,

    [Parameter()]
    [switch]$Verify
)

begin {
    # Logging-Konfiguration
    $logFile = Join-Path $PSScriptRoot "ADUserCreation_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    
    function Write-Log {
        param(
            [string]$Message,
            [ValidateSet('INFO','WARNING','ERROR','SUCCESS')]
            [string]$Level = 'INFO'
        )
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "[$timestamp][$Level] $Message"
        Add-Content -Path $logFile -Value $logEntry
        
        # Farbauswahl
        $color = switch ($Level) {
            'INFO'    { 'Gray' }
            'WARNING' { 'Yellow' }
            'ERROR'   { 'Red' }
            'SUCCESS' { 'Green' }
        }
        
        Write-Host ("{0}: {1}" -f $Level.PadRight(7), $Message) -ForegroundColor $color
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
            Write-Log "Gruppenmitgliedschaften kopiert" -Level SUCCESS
        }
        catch {
            Write-Log "Fehler beim Gruppenkopieren: $_" -Level ERROR
            throw
        }
    }
}

process {
    try {
        # Vorlagenbenutzer überprüfen
        $template = Get-ADUser -Identity $TemplateUser -Properties *
        if (-not $template) {
            throw "Vorlagenbenutzer $TemplateUser nicht gefunden"
        }

        # Duplikatsprüfung
        if (Get-ADUser -Filter "SamAccountName -eq '$NewUserName'") {
            throw "Benutzer $NewUserName existiert bereits"
        }

        # Passwortvalidierung
        if ($NewUserPassword.Length -lt 12 -or 
            -not ($NewUserPassword -match '\d') -or
            -not ($NewUserPassword -match '[A-Z]') -or
            -not ($NewUserPassword -match '[a-z]')) {
            throw "Passwort entspricht nicht den Sicherheitsrichtlinien"
        }

        # Benutzer erstellen
        $newUserParams = @{
            Instance              = $template
            SamAccountName        = $NewUserName
            UserPrincipalName     = "$NewUserName@$((Get-ADDomain).DNSRoot)"
            Name                  = $NewUserName
            AccountPassword       = ConvertTo-SecureString $NewUserPassword -AsPlainText -Force
            Enabled               = $true
            ChangePasswordAtLogon = $true
        }

        if ($PSCmdlet.ShouldProcess($NewUserName, "AD-Benutzer erstellen")) {
            $newUser = New-ADUser @newUserParams -PassThru
            Copy-ADGroupMembership -SourceUser $TemplateUser -TargetUser $NewUserName
            Write-Log "Benutzer $NewUserName erfolgreich erstellt" -Level SUCCESS
        }

        # Verifikation
        if ($Verify) {
            $newUser = Get-ADUser -Identity $NewUserName -Properties *
            $comparisonProps = @(
                'Department', 'Title', 'Office', 
                'Company', 'Manager', 'Enabled'
            )

            foreach ($prop in $comparisonProps) {
                if ($newUser.$prop -ne $template.$prop) {
                    Write-Log "Abweichung bei $prop: Vorlage[$($template.$prop)] Neu[$($newUser.$prop)]" -Level WARNING
                }
            }
        }
    }
    catch {
        Write-Log "FEHLER: $_" -Level ERROR
        exit 1
    }
}

end {
    Write-Log "Prozess abgeschlossen. Log-Datei: $logFile" -Level INFO
}
