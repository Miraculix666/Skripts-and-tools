<#
╔═════════════════════════════════════════════════════════════════
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
╚═════════════════════════════════════════════════════════════════

# ... Dateiinhalt wie im Original, aber mit den Korrekturen ...
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
$script:CsvEncoding = 'utf8'  # Korrektur: Standard PowerShell Wert ohne BOM

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
    param()

    try {
        if (-not (Test-Path $script:LogDir -PathType Container)) {
            New-Item -Path $script:LogDir -ItemType Directory -Force | Out-Null
            Write-Verbose "Log-Verzeichnis erstellt: $($script:LogDir)"
        }

        $header = @"
╔═════════════════════════════════════════════════════════════════
║                   MANAGE-RDPUSERS.PS1 - AUSFÜHRUNGS-LOG (v3.0)                 ║
╚═════════════════════════════════════════════════════════════════

Skript-Start:              $($script:GlobalDateTime)
Modus:                     $($PSCmdlet.ParameterSetName)
Ausführender Benutzer:     $($env:USERNAME)@$($env:COMPUTERNAME)
PowerShell-Version:        $($PSVersionTable.PSVersion)
Skript-Verzeichnis:        $($script:ScriptRoot)
Log-Datei:                 $($script:LogFile)

──────────────────────────────────────────────────────────────────
PARAMETER
──────────────────────────────────────────────────────────────────
UserListPath:              $UserListPath
ClientListPath:            $ClientListPath
OutputPath:                $(if ($OutputPath) { $OutputPath } else { '[Auto-Generated]' })
SendMethod:                $(if ($SendMethod) { $SendMethod } else { '[N/A]' })
SmtpServer:                $(if ($SmtpServer) { $SmtpServer } else { '[N/A]' })
SendTime:                  $(if ($SendTime) { $SendTime } else { '[N/A]' })
WhatIf Mode:               $(if ($WhatIfPreference) { 'JA' } else { 'NEIN' })

──────────────────────────────────────────────────────────────────
AUSFÜHRUNGS-PROTOKOLL
──────────────────────────────────────────────────────────────────

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
    param()

    $footer = @"

──────────────────────────────────────────────────────────────────
AUSFÜHRUNGS-ZUSAMMENFASSUNG
──────────────────────────────────────────────────────────────────

Erfolgreich verarbeitete Einträge:        $($script:ExecutionLog.Count)
Warnungen:                                $($script:WarningLog.Count)
Fehler:                                   $($script:ErrorLog.Count)

Skript-Ende:                              $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')

──────────────────────────────────────────────────────────────────
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
# DATEN-IMPORT & VALIDIERUNG
#==============================================================================

Function Import-CsvData {
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
        $recordCount = ($data | Measure-Object).Count
        Write-LogMessage -Message "✓ $recordCount Einträge geladen aus: $(Split-Path $Path -Leaf)" -Level 'SUCCESS' -Color $script:ColorSuccess
        if ($ColumnName -and $recordCount -gt 0) {
            if (-not ($data[0].PSObject.Properties.Name -contains $ColumnName)) {
                throw "Erforderliche Spalte nicht gefunden: '$ColumnName'"
            }
        }
        return $data
    }
    catch {
        $msg = if ($_.Exception) { $_.Exception.Message } else { $_.ToString() }
        Write-LogMessage -Message "✗ Fehler beim Laden der CSV: $msg" -Level 'ERROR' -Color $script:ColorError
        return $null
    }
}

#==============================================================================
# HILFSFUNKTIONEN - FILE OPERATIONS
#==============================================================================

Function New-RemoteDesktopFile {
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
        $dir = Split-Path $FilePath
        if (!(Test-Path $dir)) { 
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
        Set-Content -Path $FilePath -Value $rdpContent -Encoding $script:CsvEncoding -ErrorAction Stop
        if (Test-Path $FilePath) {
            Write-LogMessage -Message "RDP-Datei erstellt: $FilePath" -Level 'SUCCESS' -Color $script:ColorSuccess
            return $true
        } else {
            throw "RDP-Datei konnte nicht geschrieben werden."
        }
    }
    catch {
        $msg = if ($_.Exception) { $_.Exception.Message } else { $_.ToString() }
        Write-LogMessage -Message "Fehler beim Erstellen der RDP-Datei: $msg" -Level 'ERROR' -Color $script:ColorError
        return $false
    }
}

#==============================================================================
# HILFSFUNKTIONEN - EMAIL OPERATIONS
#==============================================================================

Function New-OutlookMailMessage {
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
        } else {
            Write-LogMessage -Message "⚠ RDP-Anhang nicht gefunden: $RDPFilePath" -Level 'WARN' -Color $script:ColorWarning
        }
        $mail.SaveAs($MSGSavePath, 5)
        Write-LogMessage -Message "MSG-Datei erstellt: $MSGSavePath" -Level 'SUCCESS' -Color $script:ColorSuccess
        return $true
    }
    catch {
        $msg = if ($_.Exception) { $_.Exception.Message } else { $_.ToString() }
        Write-LogMessage -Message "Fehler beim Erstellen der MSG-Datei: $msg" -Level 'ERROR' -Color $script:ColorError
        return $false
    }
    finally {
        if ($mail) { 
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($mail) | Out-Null } catch {}
        }
        if ($outlook) {
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null } catch {}
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

#==============================================================================
# HILFSFUNKTIONEN - REMOTE OPERATIONS
#==============================================================================

Function Invoke-RemoteGroupMembership {
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
        $errorMsg = if ($_.Exception) { $_.Exception.Message.Trim() } else { $_.ToString().Trim() }
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
# HILFSFUNKTIONEN - EMAIL OPERATIONS (SMTP & Exchange scheduled)
#==============================================================================

Function Send-EmailViaSMTP {
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
            Encoding   = $script:CsvEncoding
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
        $msg = if ($_.Exception) { $_.Exception.Message } else { $_.ToString() }
        Write-LogMessage -Message "✗ SMTP-Fehler: $msg" -Level 'ERROR' -Color $script:ColorError
        return $false
    }
}

Function Send-EmailViaExchangeScheduled {
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
        $sendDateTime = [DateTime]::Parse($SendTime)
        $delayedDeliveryTime = $sendDateTime.ToString("dd MMM yyyy HH:mm:ss zzz", [System.Globalization.CultureInfo]::InvariantCulture)
        Write-Verbose "Verzögerter Versand geplant für: $delayedDeliveryTime"
        $mailMessage = New-Object System.Net.Mail.MailMessage
        $mailMessage.From = $From
        $mailMessage.To.Add($To)
        $mailMessage.Subject = $Subject
        $mailMessage.Body = $BodyHTML
        $mailMessage.IsBodyHtml = $true
        if (Test-Path $AttachmentPath) {
            $attachment = New-Object System.Net.Mail.Attachment($AttachmentPath)
            $mailMessage.Attachments.Add($attachment)
        }
        $mailMessage.Headers.Add("X-MS-Exchange-Organization-Delayed-Delivery", $delayedDeliveryTime)
        $mailMessage.Headers.Add("X-Delayed-Delivery", $delayedDeliveryTime)
        $smtpClient = New-Object System.Net.Mail.SmtpClient($SmtpServer, 587)
        $smtpClient.EnableSsl = $true
        if ($Credential) {
            $smtpClient.Credentials = $Credential.GetNetworkCredential()
        }
        $smtpClient.Send($mailMessage)
        Write-LogMessage -Message "✓ E-Mail an Exchange übergeben (Versand: $SendTime) für: $To" -Level 'SUCCESS' -Color $script:ColorSuccess
        # Cleanup
        $mailMessage.Dispose()
        $smtpClient.Dispose()
        return $true
    }
    catch {
        $msg = if ($_.Exception) { $_.Exception.Message } else { $_.ToString() }
        Write-LogMessage -Message "✗ Exchange-Fehler: $msg" -Level 'ERROR' -Color $script:ColorError
        return $false
    }
}

# Die übrigen Funktionen und der Workflow-Block sind gleich wie im Original.

#==============================================================================
# ENDE
#==============================================================================
