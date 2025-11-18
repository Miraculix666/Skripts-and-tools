<#
╔════════════════════════════════════════════════════════════════════════════════╗
║                   MANAGE-RDPUSERS.PS1 (v3.0 - EMAIL-EDITION)                   ║
║                                                                                ║
║  PowerShell 5.1 Script für Remotedesktop-Benutzerverwaltung + E-Mail-Versand   ║
║  Mit 3 E-Mail-Optionen: Outlook MSG, SMTP direkt, Exchange scheduled           ║
║                                                                                ║
║  Autor: PS-Coding (via KI-Optimierung)                                         ║
║  Version: 3.0 (E-Mail-Edition)                                                 ║
║  Erstellt: 18.11.2025                                                          ║
║  PowerShell: 5.1+ erforderlich                                                 ║
║  Lokalisierung: Deutsch (DE)                                                   ║
╚════════════════════════════════════════════════════════════════════════════════╝

.DESCRIPTION
    Verwaltung von Remote-Desktop-Berechtigungen + RDP-Dateien + E-Mail-Versand.
    Unterstützt drei Betriebsmodi:
    - Modus 1: RDP-Rechte für Benutzer auf Clients HINZUFÜGEN (N:M)
    - Modus 2: RDP-Rechte für Benutzer von Clients ENTFERNEN (N:M)
    - Modus 3: RDP-Dateien erstellen + E-Mail-Versand (3 Methoden)

.FEATURES
    • Drei Betriebsmodi (Rechte hinzufügen, entfernen, RDP+E-Mail)
    • Drei E-Mail-Versandmethoden:
      → Outlook MSG-Dateien (manuell versenden)
      → SMTP direkt (sofortiger Versand)
      → Exchange scheduled (verzögerter Versand, PC nicht erforderlich)
    • Automatische Planung, Bestätigung und Ausführung
    • Ping-Test zur Erreichbarkeitsprüfung vor Aktionen
    • Detaillierte Text-Logfiles mit Zeitstempel
    • Umfassende Fehlerbehandlung mit kritischen Haltepunkten
    • Verbose-Logging standardmäßig aktiviert
    • Volle -WhatIf Unterstützung (Simulation ohne Ausführung)
    • Deutsche Lokalisierung (Datumsformate, Trennzeichen, Texte)
    • ADSI/WinNT für PS 5.1 Kompatibilität

.NOTES
    Version History:
    - v1.0-1.5: Basis-Implementierung mit Passwort-Reset
    - v2.0: Vollständige Überarbeitung, Logging, Fehlerbehandlung
    - v3.0: Passwort-Reset entfernt, 3 E-Mail-Versandmethoden hinzugefügt

    Requirements:
    - PowerShell 5.1+
    - Active Directory RSAT-Tools (AD Module)
    - Admin-Rechte lokal auf Ziel-Clients
    - RPC/ADSI/WinRM zwischen Admin-PC und Clients (Firewall)
    - Optional: Outlook Desktop-Client (für MSG-Dateien)
    - Optional: SMTP-Server für E-Mail-Versand
    - Optional: Exchange Server für scheduled send

    AI-Sources:
    - ADSI-Konzept: Microsoft TechNet
    - Outlook COM: Microsoft MSDN
    - AD Cmdlets: Microsoft Learn
    - SMTP + Exchange Headers: RFC 2822, Exchange Documentation

.LINK
    https://learn.microsoft.com/en-us/powershell/module/activedirectory/
    https://learn.microsoft.com/en-us/exchange/mail-flow/

.EXAMPLE
    # MODUS 1: RDP-Rechte HINZUFÜGEN
    .\Manage-RDPUsers.ps1 -SetRDPRights -UserListPath "C:\temp\users.csv" `
        -ClientListPath "C:\temp\clients.csv" -Verbose

.EXAMPLE
    # MODUS 2: RDP-Rechte ENTFERNEN
    .\Manage-RDPUsers.ps1 -RemoveRDPRights -UserListPath "C:\temp\users.csv" `
        -ClientListPath "C:\temp\clients.csv"

.EXAMPLE
    # MODUS 3A: RDP-Dateien + Outlook MSG
    .\Manage-RDPUsers.ps1 -GenerateRDPWithEmail `
        -UserListPath "C:\temp\users.csv" `
        -ClientListPath "C:\temp\clients.csv" `
        -OutputPath "C:\Output" `
        -SendMethod Outlook

.EXAMPLE
    # MODUS 3B: RDP-Dateien + SMTP direkt
    .\Manage-RDPUsers.ps1 -GenerateRDPWithEmail `
        -UserListPath "C:\temp\users.csv" `
        -ClientListPath "C:\temp\clients.csv" `
        -OutputPath "C:\Output" `
        -SendMethod SMTP `
        -SmtpServer "smtp.firma.de" `
        -EmailFrom "rdp@firma.de" `
        -Credential (Get-Credential)

.EXAMPLE
    # MODUS 3C: RDP-Dateien + Exchange scheduled (verzögert)
    .\Manage-RDPUsers.ps1 -GenerateRDPWithEmail `
        -UserListPath "C:\temp\users.csv" `
        -ClientListPath "C:\temp\clients.csv" `
        -OutputPath "C:\Output" `
        -SendMethod Exchange `
        -SmtpServer "exchange.firma.de" `
        -SendTime "2025-11-18 15:30:00" `
        -EmailFrom "rdp@firma.de" `
        -Credential (Get-Credential)

#>

[CmdletBinding(
    SupportsShouldProcess = $true,
    ConfirmImpact = 'High',
    DefaultParameterSetName = 'SetRDPRights'
)]
param(
    # ========== PARAMETER-SET 1: RDP-Rechte HINZUFÜGEN ==========
    [Parameter(
        Mandatory = $true,
        ParameterSetName = 'SetRDPRights',
        HelpMessage = 'Schalter zum Hinzufügen von RDP-Berechtigungen'
    )]
    [Switch]
    $SetRDPRights,

    # ========== PARAMETER-SET 2: RDP-Rechte ENTFERNEN ==========
    [Parameter(
        Mandatory = $true,
        ParameterSetName = 'RemoveRDPRights',
        HelpMessage = 'Schalter zum Entfernen von RDP-Berechtigungen'
    )]
    [Switch]
    $RemoveRDPRights,

    # ========== PARAMETER-SET 3: RDP-Dateien + E-Mail ==========
    [Parameter(
        Mandatory = $true,
        ParameterSetName = 'GenerateRDPWithEmail',
        HelpMessage = 'Schalter zum Erstellen von RDP-Dateien und E-Mail-Versand'
    )]
    [Switch]
    $GenerateRDPWithEmail,

    # ========== GEMEINSAME PARAMETER (Alle Modi) ==========
    [Parameter(
        Mandatory = $true,
        ParameterSetName = 'SetRDPRights',
        HelpMessage = 'Pfad zur Benutzer-CSV-Datei (Semikolon-getrennt)'
    )]
    [Parameter(
        Mandatory = $true,
        ParameterSetName = 'RemoveRDPRights',
        HelpMessage = 'Pfad zur Benutzer-CSV-Datei (Semikolon-getrennt)'
    )]
    [Parameter(
        Mandatory = $true,
        ParameterSetName = 'GenerateRDPWithEmail',
        HelpMessage = 'Pfad zur Benutzer-CSV-Datei (Semikolon-getrennt)'
    )]
    [ValidateScript({
        if (-not (Test-Path $_ -PathType Leaf)) {
            throw "Benutzer-CSV nicht gefunden: $_"
        }
        $true
    })]
    [String]
    $UserListPath,

    [Parameter(
        Mandatory = $true,
        ParameterSetName = 'SetRDPRights',
        HelpMessage = 'Pfad zur Client-CSV-Datei (Semikolon-getrennt)'
    )]
    [Parameter(
        Mandatory = $true,
        ParameterSetName = 'RemoveRDPRights',
        HelpMessage = 'Pfad zur Client-CSV-Datei (Semikolon-getrennt)'
    )]
    [Parameter(
        Mandatory = $true,
        ParameterSetName = 'GenerateRDPWithEmail',
        HelpMessage = 'Pfad zur Client-CSV-Datei (Semikolon-getrennt)'
    )]
    [ValidateScript({
        if (-not (Test-Path $_ -PathType Leaf)) {
            throw "Client-CSV nicht gefunden: $_"
        }
        $true
    })]
    [String]
    $ClientListPath,

    # ========== OUTPUT-PFAD (Modus 3 erforderlich) ==========
    [Parameter(
        Mandatory = $false,
        ParameterSetName = 'GenerateRDPWithEmail',
        HelpMessage = 'Ausgabeordner für RDP-Dateien (wird erstellt, falls nicht vorhanden)'
    )]
    [String]
    $OutputPath,

    # ========== E-MAIL PARAMETER (Modus 3) ==========
    [Parameter(
        Mandatory = $false,
        ParameterSetName = 'GenerateRDPWithEmail',
        HelpMessage = 'E-Mail-Versandmethode: Outlook (MSG), SMTP (direkt), Exchange (scheduled)'
    )]
    [ValidateSet('Outlook', 'SMTP', 'Exchange')]
    [String]
    $SendMethod = 'Outlook',

    [Parameter(
        Mandatory = $false,
        ParameterSetName = 'GenerateRDPWithEmail',
        HelpMessage = 'SMTP-/Exchange-Server (z.B. smtp.firma.de oder exchange.firma.de)'
    )]
    [String]
    $SmtpServer,

    [Parameter(
        Mandatory = $false,
        ParameterSetName = 'GenerateRDPWithEmail',
        HelpMessage = 'Absender-Adresse für E-Mails (z.B. rdp-system@firma.de)'
    )]
    [String]
    $EmailFrom = 'rdp-system@firma.de',

    [Parameter(
        Mandatory = $false,
        ParameterSetName = 'GenerateRDPWithEmail',
        HelpMessage = 'Betreff der E-Mail'
    )]
    [String]
    $EmailSubject = 'Ihre RDP-Zugangsdaten',

    [Parameter(
        Mandatory = $false,
        ParameterSetName = 'GenerateRDPWithEmail',
        HelpMessage = 'Sendezeitpunkt für Exchange scheduled send (Format: yyyy-MM-dd HH:mm:ss)'
    )]
    [String]
    $SendTime,

    [Parameter(
        Mandatory = $false,
        ParameterSetName = 'GenerateRDPWithEmail',
        HelpMessage = 'SMTP-Anmeldedaten (Get-Credential)'
    )]
    [PSCredential]
    $Credential,

    # ========== OPTIONALE PARAMETER ==========
    [Parameter(
        Mandatory = $false,
        HelpMessage = 'Name der Spalte mit Benutzernamen in der CSV (Standard: sAMAccountName)'
    )]
    [String]
    $UserColumn = 'sAMAccountName',

    [Parameter(
        Mandatory = $false,
        HelpMessage = 'Name der Spalte mit Computernamen in der CSV (Standard: ComputerName)'
    )]
    [String]
    $ClientColumn = 'ComputerName',

    [Parameter(
        Mandatory = $false,
        HelpMessage = 'Aktuelle Logging-Stufe (Default: Verbose)'
    )]
    [ValidateSet('Silent', 'Normal', 'Verbose')]
    [String]
    $LogLevel = 'Verbose'
)

#==============================================================================
# GLOBALE VARIABLEN & KONFIGURATION
#==============================================================================

$ErrorActionPreference = 'Stop'
$WarningPreference = 'Continue'
$VerbosePreference = if ($LogLevel -eq 'Verbose') { 'Continue' } else { 'SilentlyContinue' }

# Globale Zeitstempel für konsistente Logging
$script:GlobalTimestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$script:GlobalDateTime = Get-Date -Format 'dd.MM.yyyy HH:mm:ss'

# Log- und Ausgabe-Verzeichnisse
$script:ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
$script:LogDir = Join-Path -Path $script:ScriptRoot -ChildPath 'Logs'
$script:LogFile = Join-Path -Path $script:LogDir -ChildPath "RDP_Management_$($script:GlobalTimestamp).log"

# Error-Tracking
$script:ErrorLog = New-Object System.Collections.ArrayList
$script:WarningLog = New-Object System.Collections.ArrayList
$script:ExecutionLog = New-Object System.Collections.ArrayList

# CSV-Trennzeichen (Deutsch: Semikolon)
$script:CsvDelimiter = ';'
$script:CsvEncoding = 'UTF8'

# Farben für Konsolen-Output (Zugänglichkeit)
$script:ColorSuccess = 'Green'
$script:ColorWarning = 'Yellow'
$script:ColorError = 'Red'
$script:ColorInfo = 'Cyan'
$script:ColorSection = 'DarkGray'

#==============================================================================
# LOGGING & ERROR HANDLING FRAMEWORK
#==============================================================================

Function Initialize-Logging {
    <#
    .SYNOPSIS
        Initialisiert das Logging-System und erstellt notwendige Verzeichnisse.
    #>
    param()
    
    try {
        if (-not (Test-Path $script:LogDir -PathType Container)) {
            New-Item -Path $script:LogDir -ItemType Directory -Force | Out-Null
            Write-Verbose "Log-Verzeichnis erstellt: $($script:LogDir)"
        }
        
        $header = @"
╔════════════════════════════════════════════════════════════════════════════════╗
║                   MANAGE-RDPUSERS.PS1 - AUSFÜHRUNGS-LOG (v3.0)                 ║
╚════════════════════════════════════════════════════════════════════════════════╝

Skript-Start:              $($script:GlobalDateTime)
Modus:                     $($PSCmdlet.ParameterSetName)
Ausführender Benutzer:     $($env:USERNAME)@$($env:COMPUTERNAME)
PowerShell-Version:        $($PSVersionTable.PSVersion)
Skript-Verzeichnis:        $($script:ScriptRoot)
Log-Datei:                 $($script:LogFile)

────────────────────────────────────────────────────────────────────────────────
PARAMETER
────────────────────────────────────────────────────────────────────────────────
UserListPath:              $UserListPath
ClientListPath:            $ClientListPath
OutputPath:                $(if ($OutputPath) { $OutputPath } else { '[Auto-Generated]' })
SendMethod:                $(if ($SendMethod) { $SendMethod } else { '[N/A]' })
SmtpServer:                $(if ($SmtpServer) { $SmtpServer } else { '[N/A]' })
SendTime:                  $(if ($SendTime) { $SendTime } else { '[N/A]' })
WhatIf Mode:               $(if ($WhatIfPreference) { 'JA' } else { 'NEIN' })

────────────────────────────────────────────────────────────────────────────────
AUSFÜHRUNGS-PROTOKOLL
────────────────────────────────────────────────────────────────────────────────

"@
        
        $header | Out-File -FilePath $script:LogFile -Encoding $script:CsvEncoding -Force
        Write-Verbose "Logging-System initialisiert. Log: $($script:LogFile)"
    }
    catch {
        Write-Warning "Fehler beim Initialisieren des Logging-Systems: $_"
        Write-Warning "Logging kann begrenzt sein."
    }
}

Function Write-LogMessage {
    <#
    .SYNOPSIS
        Schreibt eine Nachricht in die Log-Datei und Console.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [String]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS', 'DEBUG')]
        [String]$Level = 'INFO',

        [Parameter(Mandatory = $false)]
        [String]$Color = $null
    )

    $timestamp = Get-Date -Format 'HH:mm:ss'
    $logEntry = "[$timestamp] [$Level] $Message"

    try {
        Add-Content -Path $script:LogFile -Value $logEntry -Encoding $script:CsvEncoding -ErrorAction SilentlyContinue
    }
    catch {
        # Fehler beim Schreiben in Log wird stillschweigend ignoriert
    }

    if ($null -ne $Color -and -not $WhatIfPreference) {
        Write-Host $Message -ForegroundColor $Color
    }
    elseif (-not $WhatIfPreference) {
        Write-Host $Message
    }

    switch ($Level) {
        'ERROR' { $script:ErrorLog.Add($Message) | Out-Null }
        'WARN' { $script:WarningLog.Add($Message) | Out-Null }
        'SUCCESS' { $script:ExecutionLog.Add("✓ $Message") | Out-Null }
        'INFO' { $script:ExecutionLog.Add("• $Message") | Out-Null }
    }
}

Function Finalize-Logging {
    <#
    .SYNOPSIS
        Zeigt Zusammenfassung der Ausführung und speichert finale Log-Einträge.
    #>
    param()

    $footer = @"

────────────────────────────────────────────────────────────────────────────────
AUSFÜHRUNGS-ZUSAMMENFASSUNG
────────────────────────────────────────────────────────────────────────────────

Erfolgreich verarbeitete Einträge:        $($script:ExecutionLog.Count)
Warnungen:                                $($script:WarningLog.Count)
Fehler:                                   $($script:ErrorLog.Count)

Skript-Ende:                              $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')

────────────────────────────────────────────────────────────────────────────────
"@

    if ($script:ErrorLog.Count -gt 0) {
        $footer += @"
FEHLER-ZUSAMMENFASSUNG:
$(($script:ErrorLog | ForEach-Object { "  ✗ $_" }) -join "`n")

"@
    }

    if ($script:WarningLog.Count -gt 0) {
        $footer += @"
WARNUNGEN-ZUSAMMENFASSUNG:
$(($script:WarningLog | ForEach-Object { "  ⚠ $_" }) -join "`n")

"@
    }

    $footer += "═" * 80 + "`n"

    Add-Content -Path $script:LogFile -Value $footer -Encoding $script:CsvEncoding -ErrorAction SilentlyContinue

    Write-Host ""
    Write-Host "─" * 80 -ForegroundColor $script:ColorSection
    Write-Host "AUSFÜHRUNGS-ZUSAMMENFASSUNG" -ForegroundColor $script:ColorInfo
    Write-Host "─" * 80 -ForegroundColor $script:ColorSection
    Write-Host "Erfolgreich verarbeitete Einträge:     $($script:ExecutionLog.Count)" -ForegroundColor $script:ColorSuccess
    Write-Host "Warnungen:                             $($script:WarningLog.Count)" -ForegroundColor $script:ColorWarning
    Write-Host "Fehler:                                $($script:ErrorLog.Count)" -ForegroundColor $(if ($script:ErrorLog.Count -gt 0) { $script:ColorError } else { $script:ColorSuccess })
    Write-Host ""
    Write-Host "Log-Datei gespeichert unter:"
    Write-Host "  $($script:LogFile)" -ForegroundColor $script:ColorInfo
    Write-Host "─" * 80 -ForegroundColor $script:ColorSection
    Write-Host ""
}

#==============================================================================
# HILFSFUNKTIONEN - ACTIVE DIRECTORY
#==============================================================================

Function Get-DynamicADDomain {
    <#
    .SYNOPSIS
        Ermittelt dynamisch die AD-Domäne (NetBIOS und FQDN).
    #>
    param()

    Write-LogMessage -Message "Ermittle AD-Domänenkontext..." -Level 'INFO'

    try {
        $adDomain = Get-ADDomain -ErrorAction Stop
        $netBIOS = $adDomain.NetBIOSName
        $fqdn = $adDomain.DNSRoot

        Write-LogMessage -Message "AD-Domäne erkannt: $netBIOS ($fqdn)" -Level 'SUCCESS' -Color $script:ColorSuccess

        return [PSCustomObject]@{
            NetBIOS = $netBIOS
            FQDN    = $fqdn
            Status  = 'OK'
        }
    }
    catch {
        Write-LogMessage -Message "Fehler beim Ermitteln der AD-Domäne: $($_.Exception.Message)" -Level 'WARN' -Color $script:ColorWarning
        Write-LogMessage -Message "Verwende Fallback 'WORKGROUP' für lokale Konten." -Level 'WARN' -Color $script:ColorWarning

        return [PSCustomObject]@{
            NetBIOS = "WORKGROUP"
            FQDN    = $null
            Status  = 'FALLBACK'
        }
    }
}

Function Get-UserEmailAddress {
    <#
    .SYNOPSIS
        Holt E-Mail-Adresse eines Benutzers aus Active Directory.
    
    .PARAMETER UserName
        sAMAccountName des Benutzers
    
    .OUTPUTS
        E-Mail-Adresse oder $null bei Fehler
    #>
    param(
        [Parameter(Mandatory = $true)]
        [String]$UserName
    )

    try {
        $user = Get-ADUser -Identity $UserName -Properties EmailAddress -ErrorAction Stop
        
        if ([String]::IsNullOrWhiteSpace($user.EmailAddress)) {
            Write-LogMessage -Message "⚠ Benutzer $UserName hat keine E-Mail-Adresse im AD" -Level 'WARN' -Color $script:ColorWarning
            return $null
        }

        Write-Verbose "E-Mail-Adresse für $UserName: $($user.EmailAddress)"
        return $user.EmailAddress
    }
    catch {
        Write-LogMessage -Message "Fehler beim Abrufen der E-Mail für $UserName: $($_.Exception.Message)" -Level 'ERROR' -Color $script:ColorError
        return $null
    }
}

#==============================================================================
# HILFSFUNKTIONEN - REMOTE OPERATIONS
#==============================================================================

Function Invoke-RemoteGroupMembership {
    <#
    .SYNOPSIS
        Ändert Gruppenmitgliedschaften auf Remote-Clients via ADSI/WinNT.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [String]$ComputerName,

        [Parameter(Mandatory = $true)]
        [String]$UserName,

        [Parameter(Mandatory = $true)]
        [String]$Domain,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Add', 'Remove')]
        [String]$Action,

        [Parameter(Mandatory = $false)]
        [String]$LocalGroupName = 'Remotedesktopbenutzer'
    )

    Write-Verbose "[$ComputerName] Aktion: $Action | Benutzer: $Domain\$UserName | Gruppe: $LocalGroupName"

    try {
        $groupPath = "WinNT://$ComputerName/$LocalGroupName,group"
        $group = [ADSI]$groupPath
        $userPath = "WinNT://$Domain/$UserName,user"

        switch ($Action) {
            'Add' {
                Write-Verbose "[$ComputerName] Füge $Domain\$UserName hinzu..."
                $group.Add($userPath)
                Write-LogMessage -Message "[$ComputerName] ✓ $Domain\$UserName zu '$LocalGroupName' hinzugefügt" -Level 'SUCCESS' -Color $script:ColorSuccess
            }
            'Remove' {
                Write-Verbose "[$ComputerName] Entferne $Domain\$UserName..."
                $group.Remove($userPath)
                Write-LogMessage -Message "[$ComputerName] ✓ $Domain\$UserName von '$LocalGroupName' entfernt" -Level 'SUCCESS' -Color $script:ColorSuccess
            }
        }

        return $true
    }
    catch {
        $errorMsg = $_.Exception.Message.Trim()

        if ($errorMsg -like '*bereits*' -or $errorMsg -like '*schon*') {
            Write-LogMessage -Message "[$ComputerName] ⚠ Benutzer $Domain\$UserName ist bereits Mitglied" -Level 'WARN' -Color $script:ColorWarning
        }
        elseif ($errorMsg -like '*nicht*Mitglied*' -or $errorMsg -like '*gehört*nicht*') {
            Write-LogMessage -Message "[$ComputerName] ⚠ Benutzer $Domain\$UserName ist nicht Mitglied" -Level 'WARN' -Color $script:ColorWarning
        }
        else {
            Write-LogMessage -Message "[$ComputerName] ✗ ADSI-Fehler: $errorMsg" -Level 'ERROR' -Color $script:ColorError
            Write-LogMessage -Message "[$ComputerName] Mögliche Ursachen: Admin-Rechte, RPC-Firewall, Benutzer nicht gefunden" -Level 'DEBUG'
        }

        return $false
    }
}

#==============================================================================
# HILFSFUNKTIONEN - FILE OPERATIONS
#==============================================================================

Function New-RemoteDesktopFile {
    <#
    .SYNOPSIS
        Erstellt eine RDP-Konfigurationsdatei für Remote-Desktop-Verbindung.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [String]$ComputerName,

        [Parameter(Mandatory = $true)]
        [String]$UserName,

        [Parameter(Mandatory = $true)]
        [String]$Domain,

        [Parameter(Mandatory = $true)]
        [String]$FilePath
    )

    Write-Verbose "Erstelle RDP-Datei: $FilePath"

    $rdpContent = @"
screen mode id:i:2
desktopwidth:i:1920
desktopheight:i:1080
session bpp:i:32
compression:i:1
keyboardhook:i:2
displayconnectionbar:i:1
disable wallpaper:i:1
allow font smoothing:i:1
allow desktop composition:i:1
disable full window drag:i:1
disable menu anims:i:1
disable themes:i:1
bitmapcachepersistenable:i:1
full address:s:$ComputerName
audiomode:i:0
redirectdrives:i:0
redirectprinters:i:0
redirectcomports:i:0
redirectsmartcards:i:0
redirectclipboard:i:1
authentication level:i:2
prompt for credentials:i:1
negotiate security layer:i:1
username:s:$Domain\$UserName
domain:s:$Domain
promptcredentialonce:i:1
"@

    try {
        Set-Content -Path $FilePath -Value $rdpContent -Encoding UTF8 -ErrorAction Stop

        if (Test-Path $FilePath) {
            Write-LogMessage -Message "RDP-Datei erstellt: $FilePath" -Level 'SUCCESS' -Color $script:ColorSuccess
            return $true
        }
        else {
            throw "RDP-Datei konnte nicht geschrieben werden."
        }
    }
    catch {
        Write-LogMessage -Message "Fehler beim Erstellen der RDP-Datei: $($_.Exception.Message)" -Level 'ERROR' -Color $script:ColorError
        return $false
    }
}

#==============================================================================
# HILFSFUNKTIONEN - EMAIL OPERATIONS
#==============================================================================

Function New-OutlookMailMessage {
    <#
    .SYNOPSIS
        Erstellt eine Outlook MSG-Datei mit RDP-Anhang.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [String]$RecipientEmail,

        [Parameter(Mandatory = $true)]
        [String]$Subject,

        [Parameter(Mandatory = $true)]
        [String]$BodyHTML,

        [Parameter(Mandatory = $true)]
        [String]$RDPFilePath,

        [Parameter(Mandatory = $true)]
        [String]$MSGSavePath
    )

    Write-Verbose "Erstelle Outlook MSG-Datei: $MSGSavePath"

    $Error.Clear()
    $outlook = $null
    $mail = $null

    try {
        try {
            $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
            Write-Verbose "Nutze bestehende Outlook-Instanz"
        }
        catch {
            Write-Verbose "Starte neue Outlook-Instanz"
            $outlook = New-Object -ComObject Outlook.Application
        }

        if (-not $outlook) {
            throw "Outlook COM-Objekt konnte nicht erstellt werden. Ist Outlook installiert?"
        }

        $mail = $outlook.CreateItem(0)
        $mail.Subject = $Subject
        $mail.HTMLBody = $BodyHTML

        if (Test-Path $RDPFilePath) {
            $attachment = $mail.Attachments.Add($RDPFilePath, 1, 1, (Split-Path $RDPFilePath -Leaf))
            Write-Verbose "Anhang hinzugefügt: $RDPFilePath"
        }
        else {
            Write-LogMessage -Message "⚠ RDP-Anhang nicht gefunden: $RDPFilePath" -Level 'WARN' -Color $script:ColorWarning
        }

        $mail.SaveAs($MSGSavePath, 5)

        Write-LogMessage -Message "MSG-Datei erstellt: $MSGSavePath" -Level 'SUCCESS' -Color $script:ColorSuccess
        return $true
    }
    catch {
        Write-LogMessage -Message "Fehler beim Erstellen der MSG-Datei: $($_.Exception.Message)" -Level 'ERROR' -Color $script:ColorError
        return $false
    }
    finally {
        if ($mail) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($mail) | Out-Null
        }
        if ($outlook) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

Function Send-EmailViaSMTP {
    <#
    .SYNOPSIS
        Sendet E-Mail direkt via SMTP.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [String]$To,

        [Parameter(Mandatory = $true)]
        [String]$From,

        [Parameter(Mandatory = $true)]
        [String]$Subject,

        [Parameter(Mandatory = $true)]
        [String]$BodyHTML,

        [Parameter(Mandatory = $true)]
        [String]$SmtpServer,

        [Parameter(Mandatory = $true)]
        [String]$AttachmentPath,

        [Parameter(Mandatory = $false)]
        [PSCredential]$Credential
    )

    Write-Verbose "Sende E-Mail via SMTP an: $To"

    try {
        $mailParams = @{
            To         = $To
            From       = $From
            Subject    = $Subject
            Body       = $BodyHTML
            BodyAsHtml = $true
            SmtpServer = $SmtpServer
            Port       = 587
            UseSsl     = $true
            Encoding   = 'UTF8'
        }

        if ($Credential) {
            $mailParams['Credential'] = $Credential
        }

        if (Test-Path $AttachmentPath) {
            $mailParams['Attachments'] = $AttachmentPath
        }

        Send-MailMessage @mailParams -ErrorAction Stop

        Write-LogMessage -Message "✓ E-Mail via SMTP versendet an: $To" -Level 'SUCCESS' -Color $script:ColorSuccess
        return $true
    }
    catch {
        Write-LogMessage -Message "✗ SMTP-Fehler: $($_.Exception.Message)" -Level 'ERROR' -Color $script:ColorError
        return $false
    }
}

Function Send-EmailViaExchangeScheduled {
    <#
    .SYNOPSIS
        Sendet E-Mail an Exchange mit verzögertem Versand (scheduled send).
        
    .DESCRIPTION
        Nutzt X-MS-Exchange-Organization-Delayed-Delivery Header für scheduled send.
        Server versendet Mail automatisch zum geplanten Zeitpunkt.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [String]$To,

        [Parameter(Mandatory = $true)]
        [String]$From,

        [Parameter(Mandatory = $true)]
        [String]$Subject,

        [Parameter(Mandatory = $true)]
        [String]$BodyHTML,

        [Parameter(Mandatory = $true)]
        [String]$SmtpServer,

        [Parameter(Mandatory = $true)]
        [String]$AttachmentPath,

        [Parameter(Mandatory = $true)]
        [String]$SendTime,

        [Parameter(Mandatory = $false)]
        [PSCredential]$Credential
    )

    Write-Verbose "Sende E-Mail an Exchange mit Sendezeitpunkt: $SendTime"

    try {
        # Parse SendTime
        $sendDateTime = [DateTime]::Parse($SendTime)
        
        # RFC 2822 Format für Exchange Header (mit Zeitzone)
        $delayedDeliveryTime = $sendDateTime.ToString("dd MMM yyyy HH:mm:ss zzz", [System.Globalization.CultureInfo]::InvariantCulture)
        
        Write-Verbose "Verzögerter Versand geplant für: $delayedDeliveryTime"

        # Erstelle MailMessage Objekt
        $mailMessage = New-Object System.Net.Mail.MailMessage
        $mailMessage.From = $From
        $mailMessage.To.Add($To)
        $mailMessage.Subject = $Subject
        $mailMessage.Body = $BodyHTML
        $mailMessage.IsBodyHtml = $true

        # Füge Anhang hinzu
        if (Test-Path $AttachmentPath) {
            $attachment = New-Object System.Net.Mail.Attachment($AttachmentPath)
            $mailMessage.Attachments.Add($attachment)
        }

        # Setze Exchange Delayed Delivery Header
        $mailMessage.Headers.Add("X-MS-Exchange-Organization-Delayed-Delivery", $delayedDeliveryTime)
        $mailMessage.Headers.Add("X-Delayed-Delivery", $delayedDeliveryTime)
        
        # SMTP-Client konfigurieren
        $smtpClient = New-Object System.Net.Mail.SmtpClient($SmtpServer, 587)
        $smtpClient.EnableSsl = $true

        if ($Credential) {
            $smtpClient.Credentials = $Credential.GetNetworkCredential()
        }

        # Sende Mail
        $smtpClient.Send($mailMessage)

        Write-LogMessage -Message "✓ E-Mail an Exchange übergeben (Versand: $SendTime) für: $To" -Level 'SUCCESS' -Color $script:ColorSuccess
        
        # Cleanup
        $mailMessage.Dispose()
        $smtpClient.Dispose()
        
        return $true
    }
    catch {
        Write-LogMessage -Message "✗ Exchange-Fehler: $($_.Exception.Message)" -Level 'ERROR' -Color $script:ColorError
        return $false
    }
}

#==============================================================================
# DATEN-IMPORT & VALIDIERUNG
#==============================================================================

Function Import-CsvData {
    <#
    .SYNOPSIS
        Lädt und validiert CSV-Datei (Benutzer oder Clients).
    #>
    param(
        [Parameter(Mandatory = $true)]
        [String]$Path,

        [Parameter(Mandatory = $false)]
        [String]$ColumnName = $null
    )

    Write-LogMessage -Message "Lade CSV-Datei: $Path" -Level 'INFO'

    try {
        $data = Import-Csv -Path $Path -Delimiter $script:CsvDelimiter -Encoding $script:CsvEncoding -ErrorAction Stop

        if ($null -eq $data) {
            throw "CSV-Datei ist leer oder konnte nicht gelesen werden."
        }

        $recordCount = if ($data -is [Array]) { $data.Count } else { 1 }
        Write-LogMessage -Message "✓ $recordCount Einträge geladen aus: $(Split-Path $Path -Leaf)" -Level 'SUCCESS' -Color $script:ColorSuccess

        if ($ColumnName -and $data -is [Array]) {
            $firstRecord = $data[0]
            if (-not ($firstRecord.PSObject.Properties.Name -contains $ColumnName)) {
                throw "Erforderliche Spalte nicht gefunden: '$ColumnName'"
            }
        }

        return $data
    }
    catch {
        Write-LogMessage -Message "✗ Fehler beim Laden der CSV: $($_.Exception.Message)" -Level 'ERROR' -Color $script:ColorError
        return $null
    }
}

#==============================================================================
# HAUPT-WORKFLOWS
#==============================================================================

Function Invoke-RDPRightsWorkflow {
    <#
    .SYNOPSIS
        Führt Modus 1 (RDP-Rechte hinzufügen) oder Modus 2 (entfernen) durch.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Add', 'Remove')]
        [String]$Mode
    )

    Write-LogMessage -Message ("Starte RDP-Rechte-Workflow (Modus: $Mode)") -Level 'INFO'

    $users = Import-CsvData -Path $UserListPath -ColumnName $UserColumn
    $clients = Import-CsvData -Path $ClientListPath -ColumnName $ClientColumn

    if ($null -eq $users -or $null -eq $clients) {
        Write-LogMessage -Message "Erforderliche Daten konnten nicht geladen werden. Abbruch." -Level 'ERROR' -Color $script:ColorError
        return
    }

    $adContext = Get-DynamicADDomain
    if ($adContext.Status -eq 'FALLBACK') {
        Write-LogMessage -Message "⚠ AD-Verbindung fehlgeschlagen. Lokale Konten können nicht genutzt werden." -Level 'WARN' -Color $script:ColorWarning
    }

    Write-Host ""
    Write-Host "═" * 80 -ForegroundColor $script:ColorSection
    Write-Host "PHASE 1: PLANUNG & VALIDIERUNG" -ForegroundColor $script:ColorInfo
    Write-Host "═" * 80 -ForegroundColor $script:ColorSection

    $plannedActions = @()
    $offlineClients = @()

    foreach ($client in $clients) {
        $clientName = $client.($ClientColumn)

        if ([String]::IsNullOrWhiteSpace($clientName)) {
            Write-LogMessage -Message "⚠ Leerer Client-Name übersprungen" -Level 'DEBUG'
            continue
        }

        Write-Verbose "Teste Ping: $clientName"
        if (-not (Test-Connection -ComputerName $clientName -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
            Write-LogMessage -Message "⚠ Client OFFLINE: $clientName (Ping fehlgeschlagen)" -Level 'WARN' -Color $script:ColorWarning
            $offlineClients += $clientName
            continue
        }

        foreach ($user in $users) {
            $userName = $user.($UserColumn)

            if ([String]::IsNullOrWhiteSpace($userName)) {
                Write-LogMessage -Message "⚠ Leerer Benutzer-Name übersprungen" -Level 'DEBUG'
                continue
            }

            $plannedActions += [PSCustomObject]@{
                Client      = $clientName
                User        = $userName
                Mode        = $Mode
                PlannedTime = Get-Date
            }
        }
    }

    Write-LogMessage -Message "$($plannedActions.Count) Aktionen geplant, $($offlineClients.Count) Clients offline" -Level 'INFO'

    Write-Host ""
    Write-Host "─" * 80 -ForegroundColor $script:ColorSection
    Write-Host "PHASE 2: BESTÄTIGUNG" -ForegroundColor $script:ColorInfo
    Write-Host "─" * 80 -ForegroundColor $script:ColorSection

    Write-Host ""
    Write-Host "Geplante Aktionen:"
    Write-Host "  • Benutzer:           $($(@($plannedActions).User | Select-Object -Unique).Count)"
    Write-Host "  • Clients (Online):   $($(@($plannedActions).Client | Select-Object -Unique).Count)"
    Write-Host "  • Modus:              $(if ($Mode -eq 'Add') { 'HINZUFÜGEN' } else { 'ENTFERNEN' })"
    Write-Host ""

    if (-not $PSCmdlet.ShouldProcess("$($plannedActions.Count) RDP-Berechtigungen", "ändern")) {
        Write-LogMessage -Message "WhatIf-Modus: Keine Änderungen durchgeführt" -Level 'INFO'
        return
    }

    $confirmation = Read-Host "Fortfahren? (J/N)"
    if ($confirmation -ne 'J') {
        Write-LogMessage -Message "Benutzer hat Bestätigung abgelehnt. Abbruch." -Level 'WARN' -Color $script:ColorWarning
        return
    }

    Write-Host ""
    Write-Host "═" * 80 -ForegroundColor $script:ColorSection
    Write-Host "PHASE 3: AUSFÜHRUNG" -ForegroundColor $script:ColorInfo
    Write-Host "═" * 80 -ForegroundColor $script:ColorSection
    Write-Host ""

    $successCount = 0
    $failCount = 0

    foreach ($action in $plannedActions) {
        $result = Invoke-RemoteGroupMembership `
            -ComputerName $action.Client `
            -UserName $action.User `
            -Domain $adContext.NetBIOS `
            -Action $Mode

        if ($result) {
            $successCount++
        }
        else {
            $failCount++
        }
    }

    Write-Host ""
    Write-Host "─" * 80 -ForegroundColor $script:ColorSection
    Write-Host "ERGEBNISSE" -ForegroundColor $script:ColorInfo
    Write-Host "─" * 80 -ForegroundColor $script:ColorSection
    Write-Host "Erfolgreich:  $successCount" -ForegroundColor $script:ColorSuccess
    Write-Host "Fehler:       $failCount" -ForegroundColor $(if ($failCount -gt 0) { $script:ColorError } else { $script:ColorSuccess })
    Write-Host ""

    Write-LogMessage -Message "RDP-Workflow abgeschlossen: $successCount erfolgreich, $failCount Fehler" -Level 'SUCCESS'
}

Function Invoke-GenerateRDPWithEmailWorkflow {
    <#
    .SYNOPSIS
        Führt Modus 3 (RDP-Dateien erstellen + E-Mail-Versand) durch.
    #>
    param()

    Write-LogMessage -Message "Starte RDP-Generierung + E-Mail-Workflow (Methode: $SendMethod)" -Level 'INFO'

    $users = Import-CsvData -Path $UserListPath -ColumnName $UserColumn
    $clients = Import-CsvData -Path $ClientListPath -ColumnName $ClientColumn

    if ($null -eq $users -or $null -eq $clients) {
        Write-LogMessage -Message "Erforderliche Daten konnten nicht geladen werden. Abbruch." -Level 'ERROR' -Color $script:ColorError
        return
    }

    # Output-Ordner sicherstellen
    if ([String]::IsNullOrWhiteSpace($OutputPath)) {
        $OutputPath = Join-Path -Path $script:ScriptRoot -ChildPath "RDP_Ausgabe_$($script:GlobalTimestamp)"
        Write-LogMessage -Message "Auto-Generierter Output-Ordner: $OutputPath" -Level 'INFO'
    }

    if (-not (Test-Path $OutputPath -PathType Container)) {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
        Write-LogMessage -Message "Output-Ordner erstellt: $OutputPath" -Level 'SUCCESS' -Color $script:ColorSuccess
    }

    $adContext = Get-DynamicADDomain

    # Validiere E-Mail-Parameter
    if ($SendMethod -eq 'SMTP' -or $SendMethod -eq 'Exchange') {
        if ([String]::IsNullOrWhiteSpace($SmtpServer)) {
            Write-LogMessage -Message "✗ SMTP/Exchange-Server nicht angegeben. Verwenden Sie -SmtpServer Parameter." -Level 'ERROR' -Color $script:ColorError
            return
        }
        if (-not $Credential) {
            Write-LogMessage -Message "✗ Credential nicht angegeben. Verwenden Sie -Credential (Get-Credential)." -Level 'ERROR' -Color $script:ColorError
            return
        }
    }

    if ($SendMethod -eq 'Exchange' -and [String]::IsNullOrWhiteSpace($SendTime)) {
        Write-LogMessage -Message "✗ SendTime nicht angegeben für Exchange scheduled send. Verwenden Sie -SendTime 'yyyy-MM-dd HH:mm:ss'." -Level 'ERROR' -Color $script:ColorError
        return
    }

    Write-Host ""
    Write-Host "═" * 80 -ForegroundColor $script:ColorSection
    Write-Host "PHASE 1: PLANUNG & VALIDIERUNG" -ForegroundColor $script:ColorInfo
    Write-Host "═" * 80 -ForegroundColor $script:ColorSection

    $plannedActions = @()
    $offlineClients = @()

    foreach ($client in $clients) {
        $clientName = $client.($ClientColumn)

        if ([String]::IsNullOrWhiteSpace($clientName)) {
            Write-LogMessage -Message "⚠ Leerer Client-Name übersprungen" -Level 'DEBUG'
            continue
        }

        Write-Verbose "Teste Ping: $clientName"
        if (-not (Test-Connection -ComputerName $clientName -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
            Write-LogMessage -Message "⚠ Client OFFLINE: $clientName (Ping fehlgeschlagen)" -Level 'WARN' -Color $script:ColorWarning
            $offlineClients += $clientName
            continue
        }

        foreach ($user in $users) {
            $userName = $user.($UserColumn)

            if ([String]::IsNullOrWhiteSpace($userName)) {
                Write-LogMessage -Message "⚠ Leerer Benutzer-Name übersprungen" -Level 'DEBUG'
                continue
            }

            $plannedActions += [PSCustomObject]@{
                Client = $clientName
                User   = $userName
            }
        }
    }

    Write-LogMessage -Message "$($plannedActions.Count) RDP-Dateien geplant (N:M), $($offlineClients.Count) Clients offline" -Level 'INFO'

    Write-Host ""
    Write-Host "─" * 80 -ForegroundColor $script:ColorSection
    Write-Host "PHASE 2: BESTÄTIGUNG" -ForegroundColor $script:ColorInfo
    Write-Host "─" * 80 -ForegroundColor $script:ColorSection

    Write-Host ""
    Write-Host "Geplante Workflows:"
    Write-Host "  • RDP-Dateien:        $($plannedActions.Count)"
    Write-Host "  • E-Mail-Versand:     $SendMethod"
    if ($SendMethod -eq 'Exchange') {
        Write-Host "  • Sendezeitpunkt:     $SendTime"
    }
    Write-Host "  • Ausgabe-Ordner:     $OutputPath"
    Write-Host ""

    if (-not $PSCmdlet.ShouldProcess("$($plannedActions.Count) RDP-Workflows (+ E-Mail)", "ausführen")) {
        Write-LogMessage -Message "WhatIf-Modus: Keine Änderungen durchgeführt" -Level 'INFO'
        return
    }

    $confirmation = Read-Host "Fortfahren? (J/N)"
    if ($confirmation -ne 'J') {
        Write-LogMessage -Message "Benutzer hat Bestätigung abgelehnt. Abbruch." -Level 'WARN' -Color $script:ColorWarning
        return
    }

    Write-Host ""
    Write-Host "═" * 80 -ForegroundColor $script:ColorSection
    Write-Host "PHASE 3: AUSFÜHRUNG" -ForegroundColor $script:ColorInfo
    Write-Host "═" * 80 -ForegroundColor $script:ColorSection
    Write-Host ""

    $successCount = 0
    $failCount = 0

    foreach ($action in $plannedActions) {
        $userName = $action.User
        $clientName = $action.Client

        Write-LogMessage -Message "► Verarbeite: $userName ↔ $clientName" -Level 'INFO' -Color $script:ColorInfo

        try {
            # 1. Erstelle RDP-Datei
            $safeUserName = $userName -replace '[\\/:*?"<>|]', '_'
            $rdpFileName = "${safeUserName}_${clientName}.rdp"
            $rdpFullPath = Join-Path -Path $OutputPath -ChildPath $rdpFileName

            $rdpSuccess = New-RemoteDesktopFile -ComputerName $clientName -UserName $userName `
                -Domain $adContext.NetBIOS -FilePath $rdpFullPath

            if (-not $rdpSuccess) {
                $failCount++
                continue
            }

            # 2. Hole E-Mail-Adresse
            $userEmail = Get-UserEmailAddress -UserName $userName

            if ([String]::IsNullOrWhiteSpace($userEmail)) {
                Write-LogMessage -Message "⚠ Keine E-Mail-Adresse für $userName - E-Mail wird übersprungen" -Level 'WARN' -Color $script:ColorWarning
                $failCount++
                continue
            }

            # 3. E-Mail-Body erstellen
            $emailBody = @"
<html>
<head>
    <meta charset="utf-8">
</head>
<body style="font-family:Calibri, Arial, sans-serif; font-size:11pt; line-height: 1.5;">
    <p>Hallo,</p>
    <p>anbei erhalten Sie Ihre RDP-Zugangsdaten für den Remote-Desktop-Zugriff.</p>
    
    <table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse; margin: 15px 0;">
        <tr style="background-color:#e0e0e0;">
            <td><b>Computer:</b></td>
            <td>$clientName</td>
        </tr>
        <tr>
            <td><b>Benutzername:</b></td>
            <td style="font-family:Consolas, monospace;">$userName</td>
        </tr>
    </table>
    
    <p><b>Verbindungsanleitung:</b></p>
    <ol>
        <li>Speichern Sie die angehängte <b>.rdp-Datei</b> auf Ihrem Desktop.</li>
        <li>Öffnen Sie die Datei per Doppelklick.</li>
        <li>Geben Sie Ihr Kennwort ein, wenn Sie aufgefordert werden.</li>
        <li>Akzeptieren Sie das Zertifikat beim ersten Verbindungsaufbau.</li>
    </ol>
    
    <p style="color:#666; font-size:9pt; margin-top: 20px;">
        <i>Dies ist eine automatisch generierte Nachricht vom RDP-Management-System.</i>
    </p>
</body>
</html>
"@

            # 4. E-Mail senden (je nach Methode)
            $emailSuccess = $false

            switch ($SendMethod) {
                'Outlook' {
                    $msgFileName = "${safeUserName}_${clientName}.msg"
                    $msgFullPath = Join-Path -Path $OutputPath -ChildPath $msgFileName

                    $emailSuccess = New-OutlookMailMessage -RecipientEmail $userEmail -Subject $EmailSubject `
                        -BodyHTML $emailBody -RDPFilePath $rdpFullPath -MSGSavePath $msgFullPath
                }
                'SMTP' {
                    $emailSuccess = Send-EmailViaSMTP -To $userEmail -From $EmailFrom -Subject $EmailSubject `
                        -BodyHTML $emailBody -SmtpServer $SmtpServer -AttachmentPath $rdpFullPath -Credential $Credential
                }
                'Exchange' {
                    $emailSuccess = Send-EmailViaExchangeScheduled -To $userEmail -From $EmailFrom -Subject $EmailSubject `
                        -BodyHTML $emailBody -SmtpServer $SmtpServer -AttachmentPath $rdpFullPath `
                        -SendTime $SendTime -Credential $Credential
                }
            }

            if ($emailSuccess) {
                $successCount++
            }
            else {
                $failCount++
            }
        }
        catch {
            Write-LogMessage -Message "✗ Fehler bei Verarbeitung: $($_.Exception.Message)" -Level 'ERROR' -Color $script:ColorError
            $failCount++
        }
    }

    Write-Host ""
    Write-Host "─" * 80 -ForegroundColor $script:ColorSection
    Write-Host "ERGEBNISSE" -ForegroundColor $script:ColorInfo
    Write-Host "─" * 80 -ForegroundColor $script:ColorSection
    Write-Host "Erfolgreich:  $successCount" -ForegroundColor $script:ColorSuccess
    Write-Host "Fehler:       $failCount" -ForegroundColor $(if ($failCount -gt 0) { $script:ColorError } else { $script:ColorSuccess })
    Write-Host "Ausgabe:      $OutputPath" -ForegroundColor $script:ColorInfo
    Write-Host ""

    Write-LogMessage -Message "RDP+E-Mail-Workflow abgeschlossen: $successCount erfolgreich, $failCount Fehler" -Level 'SUCCESS'
}

#==============================================================================
# MAIN SCRIPT ENTRY POINT
#==============================================================================

Initialize-Logging

try {
    Write-Verbose "Prüfe Active Directory Modul..."
    if (-not (Get-Module -Name ActiveDirectory)) {
        try {
            Import-Module ActiveDirectory -ErrorAction Stop
            Write-LogMessage -Message "✓ Active Directory Modul geladen" -Level 'SUCCESS' -Color $script:ColorSuccess
        }
        catch {
            Write-LogMessage -Message "✗ Active Directory Modul konnte nicht geladen werden. RSAT erforderlich." -Level 'ERROR' -Color $script:ColorError
            Write-LogMessage -Message "  Installieren Sie Windows RSAT-Tools für Active Directory" -Level 'ERROR'
            exit 1
        }
    }

    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════════════════════════════════╗" -ForegroundColor $script:ColorInfo
    Write-Host "║                  RDP USER MANAGEMENT - v3.0 (E-MAIL-EDITION)                   ║" -ForegroundColor $script:ColorInfo
    Write-Host "╚════════════════════════════════════════════════════════════════════════════════╝" -ForegroundColor $script:ColorInfo
    Write-Host ""

    switch ($PSCmdlet.ParameterSetName) {
        'SetRDPRights' {
            Invoke-RDPRightsWorkflow -Mode 'Add'
        }
        'RemoveRDPRights' {
            Invoke-RDPRightsWorkflow -Mode 'Remove'
        }
        'GenerateRDPWithEmail' {
            Invoke-GenerateRDPWithEmailWorkflow
        }
    }
}
catch {
    Write-LogMessage -Message "Kritischer Fehler: $($_.Exception.Message)" -Level 'ERROR' -Color $script:ColorError
    exit 1
}
finally {
    Finalize-Logging

    Write-Host ""
    Write-Host "Script-Ausführung abgeschlossen." -ForegroundColor $script:ColorSuccess
    Write-Host ""
}
