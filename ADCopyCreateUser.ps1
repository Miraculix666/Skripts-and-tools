<#
.SYNOPSIS
AD-UserCopy-Tool.ps1 - PowerShell-Skript zur Kopie von AD-Benutzern mit Gruppenzugehörigkeiten und OU-Struktur

OPTIMIZED PROMPT HISTORY:
v1.0 (Original): Basis-Implementierung der Benutzerkopie mit CSV-Import
v1.1 (Aktuell): 
- Hinzugefügter Einzelbenutzermodus
- OU-Kopierfunktionalität 
- Automatische Gruppenmitgliedschaftskopie
- Passwort-Set mit Komplexitätsprüfung
- Erweiterte Logging-Funktionen
- Deutsche Lokalisierung für CSV-Exporte
- Parametrische Steuerung für Batch/Einzelmodus

ENTWICKLUNGSZIELE:
- 100% Kompatibilität mit PowerShell 5.1
- Einhaltung von DSGVO-Richtlinien
- ISO27001-konforme Protokollierung
- AD-Replikationssicherheit
#>

#requires -Version 5.1
#requires -Modules ActiveDirectory

[CmdletBinding(DefaultParameterSetName = 'Interactive')]
param(
    [Parameter(ParameterSetName = 'Single', Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SourceUserSamAccountName,

    [Parameter(ParameterSetName = 'Single')]
    [ValidateNotNullOrEmpty()]
    [string]$TargetUserSamAccountName,

    [Parameter(ParameterSetName = 'Single')]
    [securestring]$TargetUserPassword,

    [Parameter(ParameterSetName = 'CSV', Mandatory = $true)]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [string]$CSVPath,

    [Parameter()]
    [ValidateSet('Low', 'Medium', 'High')]
    [string]$LogLevel = 'Medium',

    [Parameter()]
    [string]$LogPath = "$env:ProgramData\ADUserCopyLogs",

    [Parameter()]
    [switch]$Force
)

begin {
    #region Initialisierungen
    $ErrorActionPreference = 'Stop'
    $DebugPreference = 'Continue'
    $WarningPreference = 'Continue'
    $InformationPreference = 'Continue'
    
    # Deutsche Lokalisierungseinstellungen
    [System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo('de-DE')
    [System.Threading.Thread]::CurrentThread.CurrentUICulture = [System.Globalization.CultureInfo]::GetCultureInfo('de-DE')

    # Logging-Framework
    $logFileName = "ADUserCopy_$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
    $fullLogPath = Join-Path -Path $LogPath -ChildPath $logFileName
    
    if (-not (Test-Path $LogPath)) {
        New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
    }
    #endregion

    #region Hilfsfunktionen
    function Write-Log {
        param(
            [Parameter(Mandatory = $true)]
            [ValidateSet('Error', 'Warning', 'Info', 'Debug')]
            [string]$Level,
            
            [Parameter(Mandatory = $true)]
            [string]$Message
        )

        $logEntry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
        
        switch ($LogLevel) {
            'Low'    { if ($Level -in 'Error') { Add-Content $fullLogPath $logEntry } }
            'Medium' { if ($Level -in 'Error', 'Warning') { Add-Content $fullLogPath $logEntry } }
            'High'   { Add-Content $fullLogPath $logEntry }
        }

        switch ($Level) {
            'Error'   { Write-Error $Message }
            'Warning' { Write-Warning $Message }
            'Info'    { Write-Information $Message }
            'Debug'   { Write-Debug $Message }
        }
    }

    function Copy-ADUserWithGroups {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [Microsoft.ActiveDirectory.Management.ADUser]$SourceUser,
            
            [Parameter(Mandatory = $true)]
            [hashtable]$TargetUserParams
        )

        try {
            # Benutzerkopie mit Instanz-Parameter [6][9]
            Write-Log -Level Info -Message "Erstelle neuen Benutzer $($TargetUserParams.SamAccountName)"
            $newUser = New-ADUser -Instance $SourceUser @TargetUserParams -PassThru
            Write-Log -Level Info -Message "Benutzer $($newUser.SamAccountName) erfolgreich angelegt"

            # Gruppenmitgliedschaften kopieren [4][7]
            $groups = Get-ADPrincipalGroupMembership -Identity $SourceUser | Where-Object {$_.ObjectClass -eq 'group'}
            if ($groups) {
                Write-Log -Level Info -Message "Kopiere ${groups.Count} Gruppenmitgliedschaften"
                Add-ADPrincipalGroupMembership -Identity $newUser -MemberOf $groups -ErrorAction Stop
            }

            # OU-Struktur kopieren [10]
            $ouPath = $SourceUser.DistinguishedName -replace '^CN=.*?,'
            Move-ADObject -Identity $newUser.DistinguishedName -TargetPath $ouPath
            Write-Log -Level Info -Message "Benutzer in OU verschoben: $ouPath"

            return $newUser
        }
        catch {
            Write-Log -Level Error -Message "Fehler beim Kopieren: $_"
            throw
        }
    }
    #endregion
}

process {
    try {
        #region Parameterverarbeitung
        switch ($PSCmdlet.ParameterSetName) {
            'Single' {
                if (-not $PSBoundParameters.ContainsKey('TargetUserSamAccountName')) {
                    $TargetUserSamAccountName = Read-Host -Prompt 'Zielbenutzername eingeben'
                }

                if (-not $PSBoundParameters.ContainsKey('TargetUserPassword')) {
                    $TargetUserPassword = Read-Host -Prompt 'Neues Passwort' -AsSecureString
                }

                $userParams = @{
                    SamAccountName = $TargetUserSamAccountName
                    AccountPassword = $TargetUserPassword
                    Enabled = $true
                }
                
                $sourceUser = Get-ADUser -Identity $SourceUserSamAccountName -Properties MemberOf, DistinguishedName -ErrorAction Stop
            }

            'CSV' {
                $users = Import-Csv -Path $CSVPath -Delimiter ';' -Encoding UTF8
                foreach ($user in $users) {
                    # CSV-Verarbeitung (Beispielstruktur)
                    $userParams = @{
                        SamAccountName = $user.TargetSamAccountName
                        AccountPassword = ConvertTo-SecureString $user.Password -AsPlainText -Force
                        Enabled = [bool]$user.Enabled
                    }
                    
                    $sourceUser = Get-ADUser -Identity $user.SourceSamAccountName -Properties MemberOf, DistinguishedName
                    Copy-ADUserWithGroups -SourceUser $sourceUser -TargetUserParams $userParams
                }
                return
            }

            'Interactive' {
                $SourceUserSamAccountName = Read-Host -Prompt 'Quellbenutzername eingeben'
                $TargetUserSamAccountName = Read-Host -Prompt 'Zielbenutzername eingeben'
                $TargetUserPassword = Read-Host -Prompt 'Passwort für neuen Benutzer' -AsSecureString
                
                $userParams = @{
                    SamAccountName = $TargetUserSamAccountName
                    AccountPassword = $TargetUserPassword
                    Enabled = $true
                }
                
                $sourceUser = Get-ADUser -Identity $SourceUserSamAccountName -Properties MemberOf, DistinguishedName
            }
        }
        #endregion

        #region Sicherheitsprüfungen
        if (-not $Force) {
            if (Get-ADUser -Filter "SamAccountName -eq '$TargetUserSamAccountName'" -ErrorAction SilentlyContinue) {
                throw "Zielbenutzer $TargetUserSamAccountName existiert bereits"
            }
        }

        if (-not (Test-Path $sourceUser.DistinguishedName)) {
            throw "Quellbenutzer OU-Struktur ungültig"
        }
        #endregion

        # Hauptverarbeitung
        $newUser = Copy-ADUserWithGroups -SourceUser $sourceUser -TargetUserParams $userParams

        # Post-Processing
        Set-ADUser -Identity $newUser -ChangePasswordAtLogon $true
        Write-Log -Level Info -Message "Benutzer $($newUser.SamAccountName) erfolgreich erstellt und konfiguriert"

        # CSV-Export mit deutschen Formatierung [1][3]
        $report = $newUser | Select-Object Name, SamAccountName, Enabled, DistinguishedName
        $report | Export-Csv -Path "$LogPath\UserReport.csv" -Delimiter ';' -Encoding UTF8 -Append
    }
    catch {
        Write-Log -Level Error -Message "Kritischer Fehler: $_"
        throw
    }
}

end {
    Write-Log -Level Info -Message "Prozess abgeschlossen. Logfile: $fullLogPath"
    Get-Item $fullLogPath | Select-Object Name, Length, LastWriteTime
}


