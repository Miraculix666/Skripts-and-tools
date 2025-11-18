<#
.FILENAME Manage-RDPUsers.ps1
.DESCRIPTION
    Ein PowerShell-Skript (v5.1) zur Verwaltung von Remotedesktop-Berechtigungen 
    und zur automatisierten Erstellung von RDP-Dateien und E-Mail-Entwürfen.
    Bietet einen interaktiven Modus (ohne Parameter) und einen Parameter-Modus (für Automatisierung).
.NOTES
    Autor           : PS-Coding
    Version         : 3.3 (FIX: Kritischer ParserError 'Here-String', FIX: J/N-Bestätigungslogik, FIX: GenerateFromLog-Params, FIX: .Count-Bug)
    Erstellt am     : 16.11.2025
    PowerShell      : 5.1
    Umgebung        : Windows On-Premise (AD)
    Voraussetzungen : 1. PowerShell 5.1
                      2. Active Directory Modul (RSAT)
                      3. Administrator-Rechte (Lokal auf Clients / Domäne)
                      4. Microsoft Outlook (Desktop-Client) ODER einen SMTP-Relay-Server.
                      5. WinRM/WMI/RPC muss auf Clients erreichbar sein.

.PARAMETERSET Set-RDPRights
    Modus 1: Fügt Benutzer (N) zu Clients (M) hinzu. (Alle zu Allen)
    (Optional: -SendEmail, um Modus 3 direkt anzuschließen)

.PARAMETERSET Remove-RDPRights
    Modus 2: Entfernt Benutzer (N) von Clients (M). (Alle von Allen)
    (Optional: -SendEmail, um Modus 3 direkt anzuschließen)

.PARAMETERSET Generate-RDPFiles
    Modus 3: N:M-Workflow. Erstellt RDP-Dateien (M) und sendet E-Mails (N) mit allen RDPs.
             (Standard ist Senden. -SaveAsMsgOnly speichert nur .msg)

.PARAMETERSET GenerateFromLog
    Modus 4: Liest ein altes Log-File (von Modus 1/2) und startet den E-Mail-Workflow (Modus 3).

.EXAMPLE
    # WORKFLOW 1 (EMPFOHLEN): ADMIN-RECHTE (Modus 1) + SMTP-VERSAND (Modus 3)
    # Führen Sie dies ALS ADMINISTRATOR aus.
    $cred = Get-Credential "DOMAIN\IhrUser"
    .\Manage-RDPUsers.ps1 -SetRDPRights -UserListPath ".\users.csv" -ClientListPath ".\clients.csv" -SendEmail -SmtpServer "smtp.deine-firma.de" -Credential $cred

.EXAMPLE
    # WORKFLOW 2 (GETEILTE RECHTE / OUTLOOK):
    # Schritt 1 (Als ADMIN): Rechte setzen und 'sendMails.ps1' generieren. (J/N-Abfrage erscheint)
    .\Manage-RDPUsers.ps1 -SetRDPRights -UserListPath ".\users.csv" -ClientListPath ".\clients.csv" -SendEmail
    # (Notieren Sie sich den Pfad zur 'sendMails_[...].ps1' im Ausgabeordner)

    # Schritt 2 (Als STANDARD-BENUTZER): Führen Sie die generierte Datei in einer NEUEN, normalen Konsole aus.
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
    .\RDP_Ausgabe_2025-11-16_10-00-00\sendMails_2025-11-16_10-00-00.ps1

.EXAMPLE
    # MODUS 4 (E-MAILS AUS ALTEM LOG SENDEN):
    # (Als STANDARD-BENUTZER ausführen, um Outlook zu verwenden)
    .\Manage-RDPUsers.ps1 -GenerateFromLog -InputLogPath ".\Logs\Manage-RDPUsers_...csv"
    # ODER (mit Alias):
    .\Manage-RDPUsers.ps1 -generatemail -InputLogPath ".\Logs\Manage-RDPUsers_...csv"
#>

[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'Interactive')]
param(
    #--- MODUS 1: Schalter ---
    [Parameter(Mandatory = $false, ParameterSetName = 'Set-RDPRights')]
    [Switch]
    $SetRDPRights,

    #--- MODUS 2: Schalter ---
    [Parameter(Mandatory = $false, ParameterSetName = 'Remove-RDPRights')]
    [Switch]
    $RemoveRDPRights,

    #--- MODUS 3: Schalter ---
    [Parameter(Mandatory = $false, ParameterSetName = 'Generate-RDPFiles')]
    [Alias('generatemail')]
    [Switch]
    $GenerateRDPFiles,

    #--- MODUS 4: Schalter ---
    [Parameter(Mandatory = $false, ParameterSetName = 'GenerateFromLog')]
    [Alias('generatemail')] 
    [Switch]
    $GenerateFromLog,
    
    [Parameter(Mandatory = $true, ParameterSetName = 'GenerateFromLog')]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]
    $InputLogPath,

    #--- Gemeinsame Parameter für Modus 1, 2, 3 ---
    [Parameter(Mandatory = $true, ParameterSetName = 'Set-RDPRights')]
    [Parameter(Mandatory = $true, ParameterSetName = 'Remove-RDPRights')]
    [Parameter(Mandatory = $true, ParameterSetName = 'Generate-RDPFiles')]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]
    $UserListPath,

    [Parameter(Mandatory = $true, ParameterSetName = 'Set-RDPRights')]
    [Parameter(Mandatory = $true, ParameterSetName = 'Remove-RDPRights')]
    [Parameter(Mandatory = $true, ParameterSetName = 'Generate-RDPFiles')]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]
    $ClientListPath,

    #--- GLOBALE E-Mail-Parameter (für Modus 1, 2, 3, 4) ---
    [Parameter(Mandatory = $false)]
    [Switch]
    $SendEmail, 

    [Parameter(Mandatory = $false)]
    [Switch]
    $SaveAsMsgOnly, 

    [Parameter(Mandatory = $false)]
    [string]
    $SmtpServer, 

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]
    $Credential = $null, 

    #--- Optionale Pfad-Parameter ---
    [Parameter(Mandatory = $false)]
    [string]
    $OutputPath, 

    #--- Optionale CSV-Header-Parameter ---
    [Parameter(Mandatory = $false, ParameterSetName = 'Set-RDPRights')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Remove-RDPRights')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Generate-RDPFiles')]
    [string]
    $UserColumn = 'sAMAccountName',

    [Parameter(Mandatory = $false, ParameterSetName = 'Set-RDPRights')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Remove-RDPRights')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Generate-RDPFiles')]
    [string]
    $ClientColumn = 'ComputerName'
)

#==============================================================================
# GLOBALE VARIABLEN
#==============================================================================
$GlobalErrorLog = [System.Collections.ArrayList]::new()
$GlobalTimestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$GlobalLogDir = Join-Path -Path $PSScriptRoot -ChildPath "Logs"

#==============================================================================
# HILFSFUNKTIONEN (AD, ADSI, RDP, Outlook/SMTP)
#==============================================================================

Function Get-DynamicADDomain {
    try {
        $adDomain = Get-ADDomain -ErrorAction Stop
        $netBIOS = $adDomain.NetBIOSName
        $fqdn = $adDomain.DNSRoot
        Write-Verbose "AD-Domäne dynamisch erkannt: $netBIOS ($fqdn)"
        return [PSCustomObject]@{ NetBIOS = $netBIOS; FQDN = $fqdn }
    }
    catch {
        $msg = "Fehler beim dynamischen Ermitteln der AD-Domäne. Stelle sicher, dass das AD-Modul geladen ist."
        Write-Warning $msg
        $GlobalErrorLog.Add($msg) | Out-Null
        return [PSCustomObject]@{ NetBIOS = "WORKGROUP"; FQDN = $null }
    }
}

Function Invoke-RemoteGroupMembership {
    <#
    .SYNOPSIS (Version 2.1)
    #>
    param(
        [string]$ComputerName,
        [string]$UserName,
        [string]$Domain,
        [string]$Action, 
        [string]$LocalGroupName = "Remotedesktopbenutzer"
    )

    Write-Verbose "Aktion '$Action': User '$Domain\$UserName' -> Gruppe '$LocalGroupName' auf '$ComputerName'"
    $group = $null
    $user = $null
    $status = 'Failed' 

    try {
        $groupPath = "WinNT://$ComputerName/$LocalGroupName,group"
        $group = [ADSI]$groupPath
        $userPath = "WinNT://$Domain/$UserName,user"
        $user = [ADSI]$userPath
        
        if ($Action -eq 'Add') {
            Write-Verbose "Füge hinzu: $($user.Path) ZU $($group.Path)"
            $group.Add($user.Path)
            
            Write-Verbose "Gegenprüfung (RefreshCache und Members())..."
            $group.RefreshCache()
            $members = $group.Members() | ForEach-Object { $_.GetType().InvokeMember("Name", "GetProperty", $null, $_, $null) }
            
            if ($members -contains $UserName) {
                $status = 'Success'
            } else {
                $status = 'VerificationFailed'
            }
        }
        elseif ($Action -eq 'Remove') {
            Write-Verbose "Entferne: $($user.Path) VON $($group.Path)"
            $group.Remove($user.Path)

            Write-Verbose "Gegenprüfung (RefreshCache und Members())..."
            $group.RefreshCache()
            $members = $group.Members() | ForEach-Object { $_.GetType().InvokeMember("Name", "GetProperty", $null, $_, $null) }
            
            if (-not ($members -contains $UserName)) {
                $status = 'Success'
            } else {
                $status = 'VerificationFailed'
            }
        }
    }
    catch {
        $errorMessage = $_.Exception.Message.Trim()
        
        if ($errorMessage -like "*bereits Mitglied*") {
            $status = 'AlreadyExists'
        }
        elseif ($errorMessage -like "*nicht Mitglied*") {
            $status = 'NotMember'
        }
        else {
            $msg = ("ADSI-FEHLER bei Client {0} für User {1}: {2}" -f $ComputerName, $UserName, $errorMessage)
            Write-Warning $msg
            $GlobalErrorLog.Add($msg) | Out-Null
            $status = 'Failed'
        }
    }
    
    return $status
}

Function Create-RDPFile {
    param(
        [string]$ComputerName,
        [string]$FilePath
    )
    Write-Verbose "Erstelle RDP-Datei für $ComputerName in '$FilePath'"
    $rdpContent = @"
screen mode id:i:2
desktopwidth:i:1920
desktopheight:i:1080
full address:s:$ComputerName
prompt for credentials:i:1
redirectclipboard:i:1
"@
    try {
        Set-Content -Path $FilePath -Value $rdpContent -Encoding UTF8 -ErrorAction Stop
        return $true
    }
    catch {
        $msg = ("Fehler beim Schreiben der RDP-Datei '{0}': {1}" -f $FilePath, $_.Exception.Message)
        Write-Warning $msg
        $GlobalErrorLog.Add($msg) | Out-Null
        return $false
    }
}

Function Start-EmailWorkflow {
    <#
    .SYNOPSIS (Version 2.3)
    #>
    param(
        [string]$UserName,
        [string[]]$RDPFilePaths,
        [string]$MSGSavePath,
        [bool]$Send, 
        [string]$SmtpServer,
        [System.Management.Automation.PSCredential]$Credential
    )
    
    $adUser = $null
    try {
        if (-not (Get-Module -Name ActiveDirectory)) {
            Write-Verbose "Lade AD-Modul für E-Mail-Workflow..."
            Import-Module ActiveDirectory -ErrorAction Stop
        }
        $adUser = Get-ADUser -Identity $UserName -Properties DisplayName, EmailAddress -ErrorAction Stop
    } catch {
        $msg = "Konnte AD-Benutzer $UserName nicht finden, um E-Mail-Adresse/Namen abzurufen. Überspringe E-Mail."
        Write-Warning $msg
        if ($Global:GlobalErrorLog) { $Global:GlobalErrorLog.Add($msg) | Out-Null }
        return $false
    }
    
    $recipientName = $adUser.DisplayName
    $recipientEmail = $adUser.EmailAddress
    $senderEmail = $adUser.EmailAddress 

    if ([string]::IsNullOrWhiteSpace($recipientEmail)) {
        $msg = "Benutzer $UserName hat keine E-Mail-Adresse im AD. E-Mail kann nicht gesendet/erstellt werden."
        Write-Warning $msg
        if ($Global:GlobalErrorLog) { $Global:GlobalErrorLog.Add($msg) | Out-Null }
        return $false
    }
    
    $mailSubject = "Ihre RDP-Verbindungsdateien"
    $mailBody = @"
<html>
<body style="font-family:Calibri, Arial, sans-serif; font-size:11pt;">
<p>Hallo $recipientName,</p>
<p>anbei erhalten Sie die RDP-Verbindungsdateien für die Schulungsraum-Clients.</p>
<p>
    <b>Anleitung:</b><br>
    1. Speichern Sie die angehängten <b>.rdp-Dateien</b> auf Ihrem Desktop.<br>
    2. Öffnen Sie die Datei für den Client, mit dem Sie sich verbinden möchten.<br>
    3. Melden Sie sich mit Ihren gewohnten Anmeldeinformationen (denselben, die Sie an Ihrem Arbeitsplatz verwenden) an.
</p>
<p>
    <i>(Dies ist eine automatisch generierte Nachricht.)</i>
</p>
</body>
</html>
"@

    if (-not [string]::IsNullOrWhiteSpace($SmtpServer)) {
        # --- METHODE A: SMTP ---
        if (-not $Send) {
            Write-Warning "SMTP-Server wurde angegeben, aber Senden ist nicht aktiv (-SaveAsMsgOnly). E-Mails werden nicht via SMTP gesendet."
            return $false
        }
        
        Write-Verbose "Versende E-Mail für $UserName via SMTP ($SmtpServer)..."
        
        if ($null -eq $Credential) {
            Write-Host "SMTP-Versand erfordert Anmeldeinformationen." -ForegroundColor Yellow
            $Credential = Get-Credential -Message "Bitte SMTP-Anmeldedaten für $SmtpServer eingeben (Leer lassen für Anonym)"
        }
        
        $smtpParams = @{
            To          = $recipientEmail
            From        = $senderEmail
            Subject     = $mailSubject
            Body        = $mailBody
            BodyAsHtml  = $true
            SmtpServer  = $SmtpServer
            Attachments = $RDPFilePaths
            ErrorAction = 'Stop'
        }
        
        if ($Credential.UserName) {
            $smtpParams.Add("Credential", $Credential)
        }
        
        try {
            Send-MailMessage @smtpParams
            Write-Host " ERFOLG: E-Mail via SMTP an $recipientEmail gesendet." -ForegroundColor Cyan
            return $true
        } catch {
            $msg = ("Fehler beim Senden der SMTP-E-Mail für {0}: {1}" -f $UserName, $_.Exception.Message)
            Write-Warning $msg
            if ($Global:GlobalErrorLog) { $Global:GlobalErrorLog.Add($msg) | Out-Null }
            return $false
        }
    }
    else {
        # --- METHODE B: Outlook (COM-Objekt) ---
        Write-Verbose "Erstelle/Sende Outlook-E-Mail für $UserName..."
        $Error.Clear()
        $outlook = $null
        $mail = $null
        
        try {
            try {
                $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
                Write-Verbose "Bestehende Outlook-Instanz wird verwendet."
            } catch {
                Write-Verbose "Starte neue Outlook-Instanz..."
                $outlook = New-Object -ComObject Outlook.Application
            }

            if (!$outlook) {
                $msg = "Outlook COM-Objekt konnte nicht erstellt werden. Ist Outlook installiert?"
                Write-Warning $msg
                if ($Global:GlobalErrorLog) { $Global:GlobalErrorLog.Add($msg) | Out-Null }
                return $false
            }
            
            $mail = $outlook.CreateItem(0) 
            $mail.Subject = $mailSubject
            $mail.To = $recipientEmail
            $mail.HTMLBody = $mailBody

            foreach ($rdpPath in $RDPFilePaths) {
                if (Test-Path $rdpPath) {
                    $mail.Attachments.Add($rdpPath, 1, 1, ($rdpPath | Split-Path -Leaf))
                }
            }
            
            if (-not [string]::IsNullOrWhiteSpace($MSGSavePath)) {
                $mail.SaveAs($MSGSavePath, 5) 
                Write-Host " ERFOLG: MSG-Datei gespeichert: $MSGSavePath" -ForegroundColor Green
            }
            
            if ($Send) {
                Write-Verbose "Sende E-Mail via Outlook an $recipientEmail..."
                $mail.Send()
                Write-Host " ERFOLG: E-Mail an $recipientEmail gesendet." -ForegroundColor Cyan
            }
            return $true
        }
        catch {
            $errorMessage = $_.Exception.Message
            $msg = ("Fehler beim Erstellen/Senden der Outlook-E-Mail für {0}: {1}" -f $UserName, $errorMessage)
            Write-Warning $msg
            if ($Global:GlobalErrorLog) { $Global:GlobalErrorLog.Add($msg) | Out-Null }
            
            if ($errorMessage -like "*80080005*") {
                $msg2 = "HINWEIS: CO_E_SERVER_EXEC_FAILURE. Dies passiert oft, wenn das Skript 'Als Admin' läuft, Outlook aber als Standardbenutzer. Verwenden Sie den -SmtpServer Parameter oder führen Sie die generierte 'sendMails.ps1' als Standardbenutzer aus."
                Write-Warning $msg2
                if ($Global:GlobalErrorLog) { $Global:GlobalErrorLog.Add($msg2) | Out-Null }
            }
            return $false
        }
        finally {
            if ($mail) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($mail) | Out-Null }
            if ($outlook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null }
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }
}

#==============================================================================
# INTERAKTIVES MENÜ
#==============================================================================

Function Show-MainMenu {
    Clear-Host
    Write-Host "
   RRRRRR   DDDDDD   PPPPPP      UU   UU  SSSSS  EEEEEE  RRRRRR      MM   MM    A    NN   NN    A     GGGGGG  EEEEEE  RRRRRR
   RR   RR  DD   DD  PP   PP     UU   UU SS      EE      RR   RR     MMM MMM   A A   NNN  NN   A A   GG       EE      RR   RR
   RRRRRR   DD   DD  PPPPPP      UU   UU  SSSSS  EEEE    RRRRRR      MM M MM  AAAAA  NN N NN  AAAAA  GG   GGG EEEE    RRRRRR
   RR  RR   DD   DD  PP          UU   UU      SS EE      RR  RR      MM   MM AA   AA NN  NNN AA   AA GG    GG EE      RR  RR
   RR   RR  DDDDDD   PP           UUUUU   SSSSS  EEEEEE  RR   RR     MM   MM AA   AA NN   NN AA   AA  GGGGGG  EEEEEE  RR   RR
                    R E M O T E   U S E R   M A N A G E R (v3.3 - Syntax Fix)
" -ForegroundColor Cyan

    Write-Host "==============================================================================" -ForegroundColor Gray
    Write-Host "  Hauptmenü - Bitte wählen Sie eine Aktion:"
    Write-Host "==============================================================================" -ForegroundColor Gray
    Write-Host ""
    Write-Host "   (1) [+]" -ForegroundColor Green -NoNewline; Write-Host " Remotedesktop-Rechte " -ForegroundColor White -NoNewline; Write-Host "HINZUFÜGEN" -ForegroundColor Green
    Write-Host "   (2) [-]" -ForegroundColor Red -NoNewline; Write-Host " Remotedesktop-Rechte " -ForegroundColor White -NoNewline; Write-Host "ENTFERNEN" -ForegroundColor Red
    Write-Host "   (3) [>]" -ForegroundColor Yellow -NoNewline; Write-Host " RDP/E-Mail-Dateien " -ForegroundColor White -NoNewline; Write-Host "ERSTELLEN (aus CSV)" -ForegroundColor Yellow
    Write-Host "   (4) [>]" -ForegroundColor Cyan -NoNewline; Write-Host " E-Mails " -ForegroundColor White -NoNewline; Write-Host "ERSTELLEN (aus Log-Datei)" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "   (Q) [x]" -ForegroundColor Gray -NoNewline; Write-Host " Beenden" -ForegroundColor Gray
    Write-Host "==============================================================================" -ForegroundColor Gray
    
    $choice = Read-Host -Prompt "Ihre Wahl"
    return $choice.ToUpper()
}

Function Get-InteractiveInput {
    param($Mode)
    
    $Params = @{
        SaveAsMsgOnly = $false
        SendEmail = $false
        SmtpServer = $null
        Credential = $null
    }
    Write-Host "--- Eingabe erforderlich ---" -ForegroundColor Yellow

    if ($Mode -in ('Set-RDPRights', 'Remove-RDPRights', 'Generate-RDPFiles')) {
        while (-not (Test-Path $Params.UserListPath -PathType Leaf)) {
            $Params.UserListPath = Read-Host -Prompt "Pfad zur Benutzer-CSV (z.B. C:\temp\users.csv)"
        }
        while (-not (Test-Path $Params.ClientListPath -PathType Leaf)) {
            $Params.ClientListPath = Read-Host -Prompt "Pfad zur Client-CSV (z.B. C:\temp\clients.csv)"
        }
    }
    
    if ($Mode -eq 'GenerateFromLog') {
         while (-not (Test-Path $Params.InputLogPath -PathType Leaf)) {
            $Params.InputLogPath = Read-Host -Prompt "Pfad zur Log-Datei (z.B. C:\Daten\Scripte\Logs\Manage-RDPUsers_...csv)"
        }
    }
    
    if ($Mode -in ('Set-RDPRights', 'Remove-RDPRights') ) {
        $sendChoice = Read-Host -Prompt "E-Mails nach Abschluss direkt senden? (J/N) (Standard: N)"
        if ($sendChoice.ToUpper() -eq 'J') { $Params.SendEmail = $true }
    }
    
    if ($Mode -in ('Generate-RDPFiles', 'GenerateFromLog') -or $Params.SendEmail) {
        $Params.OutputPath = Read-Host -Prompt "Ausgabe-Ordner (Leer lassen für Auto: '.\RDP_Ausgabe_[Zeitstempel]')"
        
        $sendChoice = Read-Host -Prompt "(1) E-Mails direkt SENDEN, (2) Nur als .msg SPEICHERN (Standard: 1)"
        if ($sendChoice -eq '2') {
             $Params.SaveAsMsgOnly = $true
        } else {
             $Params.SendEmail = $true 
             $methodChoice = Read-Host -Prompt "Wie Senden? (1) via Outlook (Standard), (2) via SMTP"
             if ($methodChoice -eq '2') {
                $Params.SmtpServer = Read-Host -Prompt "SMTP-Server-Adresse"
                if (-not [string]::IsNullOrWhiteSpace($Params.SmtpServer)) {
                    $Params.Credential = Get-Credential -Message "Bitte SMTP-Anmeldedaten für $($Params.SmtpServer) eingeben (Leer lassen für Anonym)"
                }
             }
        }
    }
    
    if ($Mode -in ('Set-RDPRights', 'Remove-RDPRights', 'Generate-RDPFiles')) {
        $Params.UserColumn = Read-Host -Prompt "Name der Benutzerspalte (Standard: 'sAMAccountName')"
        if ([string]::IsNullOrWhiteSpace($Params.UserColumn)) { $Params.UserColumn = 'sAMAccountName' }
        $Params.ClientColumn = Read-Host -Prompt "Name der Client-Spalte (Standard: 'ComputerName')"
        if ([string]::IsNullOrWhiteSpace($Params.ClientColumn)) { $Params.ClientColumn = 'ComputerName' }
    }
    
    return $Params
}

#==============================================================================
# DATEN LADEN
#==============================================================================

Function Load-CsvData {
    param(
        [string]$CsvPath,
        [string]$ExpectedHeader 
    )
    
    Write-Verbose "Lade CSV-Datei: $CsvPath"
    $CsvImportParams = @{
        Delimiter = ';'
        Encoding  = 'Default' 
        ErrorAction = 'Stop'
    }

    try {
        $data = Import-Csv -Path $CsvPath @CsvImportParams
    }
    catch {
        $msg = "Fehler beim Lesen der CSV-Datei '$CsvPath'. Stelle sicher, dass das Trennzeichen ';' ist und die Datei existiert."
        Write-Error $msg
        Write-Error ("Details: {0}" -f $_.Exception.Message)
        $GlobalErrorLog.Add($msg) | Out-Null
        return $null 
    }
    
    if ($null -eq $data -or ($data | Measure-Object).Count -eq 0) {
        $msg = "CSV-Datei '$CsvPath' ist leer oder konnte nicht gelesen werden. Abbruch."
        Write-Error $msg
        $GlobalErrorLog.Add($msg) | Out-Null
        return $null
    }
    
    if (-not [string]::IsNullOrWhiteSpace($ExpectedHeader)) {
        if ($data.GetType().Name -eq 'PSCustomObject') {
             if (-not $data.PSObject.Properties.Name -contains $ExpectedHeader) {
                 $msg = "Log-Datei '$CsvPath' scheint ungültig. Erwartete Spalte '$ExpectedHeader' nicht gefunden."
                 Write-Error $msg
                 $GlobalErrorLog.Add($msg) | Out-Null
                 return $null
             }
        }
        elseif ($data.GetType().Name -eq 'Object[]') {
             if (-not $data[0].PSObject.Properties.Name -contains $ExpectedHeader) {
                 $msg = "Log-Datei '$CsvPath' scheint ungültig. Erwartete Spalte '$ExpectedHeader' nicht gefunden."
                 Write-Error $msg
                 $GlobalErrorLog.Add($msg) | Out-Null
                 return $null
             }
        }
    }

    return $data
}

#==============================================================================
# AD-Modul-Prüfung
#==============================================================================

Function Test-ADModule {
    if (-not (Get-Module -Name ActiveDirectory)) {
        try {
            Write-Verbose "Versuche, das ActiveDirectory-Modul zu importieren..."
            Import-Module ActiveDirectory -ErrorAction Stop
        } catch {
            $msg = "Das PowerShell-Modul 'ActiveDirectory' konnte nicht geladen werden (RSAT-Tools erforderlich)."
            Write-Error $msg
            $GlobalErrorLog.Add($msg) | Out-Null
            return $false
        }
    }
    return $true
}

#==============================================================================
# PROTOKOLLIERUNG
#==============================================================================

Function Write-Log {
    param(
        [object[]]$ReportData
    )
    
    $LogName = "Manage-RDPUsers"
    if ($ReportData.Length -gt 0 -and ($ReportData[0].Action -eq 'Email_Send' -or $ReportData[0].Action -eq 'Email_Save_MSG')) {
        $LogName = "Manage-RDPFiles"
    }

    try {
        if (-not (Test-Path $GlobalLogDir)) {
            New-Item -Path $GlobalLogDir -ItemType Directory -ErrorAction Stop | Out-Null
        }
        $logFile = Join-Path -Path $GlobalLogDir -ChildPath "$($LogName)_$($GlobalTimestamp).csv"
        
        $ReportData | Export-Csv -Path $logFile -NoTypeInformation -Delimiter ';' -Encoding UTF8
        Write-Host "Protokoll wurde erfolgreich gespeichert: $logFile" -ForegroundColor DarkGreen
    } catch {
        $msg = "Fehler beim Schreiben der Protokolldatei: $($_.Exception.Message)"
        Write-Warning $msg
        $GlobalErrorLog.Add($msg) | Out-Null
    }
}

Function Write-ErrorLog {
    param([string]$ErrorOutputPath)
    
    if (($GlobalErrorLog | Measure-Object).Count -gt 0) {
        try {
            if (-not (Test-Path $ErrorOutputPath)) {
                New-Item -Path $ErrorOutputPath -ItemType Directory -ErrorAction Stop | Out-Null
            }
            $errorFile = Join-Path -Path $ErrorOutputPath -ChildPath "ERROR_$($GlobalTimestamp).TXT"
            $GlobalErrorLog | Out-File -FilePath $errorFile -Encoding UTF8
            Write-Warning "Es sind Fehler aufgetreten. Details wurden in $errorFile gespeichert."
        } catch {
            Write-Warning "Konnte $errorFile nicht schreiben. Fehler: $($_.Exception.Message)"
        }
    }
}

#==============================================================================
# E-MAIL-SKRIPT GENERATOR (V3.3 - PARSER FIX)
#==============================================================================

Function Generate-SendMailsScript {
    <#
    .SYNOPSIS
        Generiert ein eigenständiges .ps1-Skript für den Outlook/SMTP-Versand.
        Dies löst das Admin-vs-Benutzer-Kontextproblem für Outlook.
    #>
    param(
        [System.Array]$UserList,
        [string[]]$RDPFilePaths,
        [string]$OutputPath,
        [string]$SmtpServer,
        [bool]$SaveAsMsgOnly
    )
    
    Write-Verbose "Generiere 'sendMails.ps1' Ausführungsskript..."
    
    $userListString = $UserList | ForEach-Object { "'$_'" } | Join-String -Separator ", "
    $rdpListString = $RDPFilePaths | ForEach-Object { $resolved = Resolve-Path $_; "'$resolved'" } | Join-String -Separator ", "
    
    $sendAction = -not $SaveAsMsgOnly
    $resolvedOutputPath = Resolve-Path $OutputPath
    
    $smtpServerString = if ([string]::IsNullOrWhiteSpace($SmtpServer)) { '$null' } else { "'$SmtpServer'" }
    $sendActionString = '$' + "$($sendAction)"

    # Definiere die Hilfsfunktionen, die das neue Skript benötigt (self-contained)
    # --- FIX V3.3: Das schließende "@" MUSS am Anfang der Zeile stehen. ---
    $scriptContent = @"
<#
.SYNOPSIS
    Dieses Skript wurde automatisch von Manage-RDPUsers.ps1 generiert.
    Es dient zum Senden von RDP-E-Mails im korrekten Benutzerkontext (z.B. als Standardbenutzer für Outlook).
    Führen Sie dieses Skript in einer normalen (Nicht-Admin) PowerShell-Konsole aus.
#>

# --- START: Erforderliche Funktionen (kopiert aus dem Hauptskript) ---

Function Start-EmailWorkflow {
    param(
        [string]$UserName,
        [string[]]$RDPFilePaths,
        [string]$MSGSavePath,
        [bool]$Send, 
        [string]$SmtpServer,
        [System.Management.Automation.PSCredential]$Credential
    )
    
    $adUser = $null
    try {
        if (-not (Get-Module -Name ActiveDirectory)) {
            Write-Verbose "Lade AD-Modul für E-Mail-Workflow..."
            Import-Module ActiveDirectory -ErrorAction Stop
        }
        $adUser = Get-ADUser -Identity $UserName -Properties DisplayName, EmailAddress -ErrorAction Stop
    } catch {
        Write-Warning "Konnte AD-Benutzer $UserName nicht finden. Überspringe E-Mail."
        return $false
    }
    
    $recipientName = $adUser.DisplayName
    $recipientEmail = $adUser.EmailAddress
    $senderEmail = $adUser.EmailAddress 

    if ([string]::IsNullOrWhiteSpace($recipientEmail)) {
        Write-Warning "Benutzer $UserName hat keine E-Mail-Adresse im AD. E-Mail kann nicht gesendet/erstellt werden."
        return $false
    }
    
    $mailSubject = "Ihre RDP-Verbindungsdateien"
    $mailBody = @"
<html>
<body style="font-family:Calibri, Arial, sans-serif; font-size:11pt;">
<p>Hallo $recipientName,</p>
<p>anbei erhalten Sie die RDP-Verbindungsdateien für die Schulungsraum-Clients.</p>
<p>
    <b>Anleitung:</b><br>
    1. Speichern Sie die angehängten <b>.rdp-Dateien</b> auf Ihrem Desktop.<br>
    2. Öffnen Sie die Datei für den Client, mit dem Sie sich verbinden möchten.<br>
    3. Melden Sie sich mit Ihren gewohnten Anmeldeinformationen (denselben, die Sie an Ihrem Arbeitsplatz verwenden) an.
</p>
<p>
    <i>(Dies ist eine automatisch generierte Nachricht.)</i>
</p>
</body>
</html>
"@

    if (-not ([string]::IsNullOrWhiteSpace($SmtpServer))) {
        # --- METHODE A: SMTP ---
        if (-not $Send) { return $false }
        Write-Verbose "Versende E-Mail für $UserName via SMTP ($SmtpServer)..."
        
        if ($null -eq $Credential) {
            Write-Host "SMTP-Versand erfordert Anmeldeinformationen." -ForegroundColor Yellow
            $Credential = Get-Credential -Message "Bitte SMTP-Anmeldedaten für $SmtpServer eingeben (Leer lassen für Anonym)"
        }
        
        $smtpParams = @{
            To          = $recipientEmail
            From        = $senderEmail
            Subject     = $mailSubject
            Body        = $mailBody
            BodyAsHtml  = $true
            SmtpServer  = $SmtpServer
            Attachments = $RDPFilePaths
            ErrorAction = 'Stop'
        }
        
        if ($Credential.UserName) {
            $smtpParams.Add("Credential", $Credential)
        }
        
        try {
            Send-MailMessage @smtpParams
            Write-Host " ERFOLG: E-Mail via SMTP an $recipientEmail gesendet." -ForegroundColor Cyan
            return $true
        } catch {
            Write-Warning ("Fehler beim Senden der SMTP-E-Mail für {0}: {1}" -f $UserName, $_.Exception.Message)
            return $false
        }
    }
    else {
        # --- METHODE B: Outlook (COM-Objekt) ---
        Write-Verbose "Erstelle/Sende Outlook-E-Mail für $UserName..."
        $Error.Clear()
        $outlook = $null
        $mail = $null
        
        try {
            try {
                $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
                Write-Verbose "Bestehende Outlook-Instanz wird verwendet."
            } catch {
                Write-Verbose "Starte neue Outlook-Instanz..."
                $outlook = New-Object -ComObject Outlook.Application
            }

            if (!$outlook) {
                Write-Warning "Outlook COM-Objekt konnte nicht erstellt werden. Ist Outlook installiert?"
                return $false
            }
            
            $mail = $outlook.CreateItem(0) 
            $mail.Subject = $mailSubject
            $mail.To = $recipientEmail
            $mail.HTMLBody = $mailBody

            foreach ($rdpPath in $RDPFilePaths) {
                if (Test-Path $rdpPath) {
                    $mail.Attachments.Add($rdpPath, 1, 1, ($rdpPath | Split-Path -Leaf))
                }
            }
            
            if (-not [string]::IsNullOrWhiteSpace($MSGSavePath)) {
                $mail.SaveAs($MSGSavePath, 5) 
                Write-Host " ERFOLG: MSG-Datei gespeichert: $MSGSavePath" -ForegroundColor Green
            }
            
            if ($Send) {
                Write-Verbose "Sende E-Mail via Outlook an $recipientEmail..."
                $mail.Send()
                Write-Host " ERFOLG: E-Mail an $recipientEmail gesendet." -ForegroundColor Cyan
            }
            return $true
        }
        catch {
            Write-Warning ("Fehler beim Erstellen/Senden der Outlook-E-Mail für {0}: {1}" -f $UserName, $_.Exception.Message)
            if ($_.Exception.Message -like "*80080005*") {
                Write-Warning "HINWEIS: CO_E_SERVER_EXEC_FAILURE. Führen Sie dieses Skript als Standardbenutzer (NICHT 'Als Admin') aus."
            }
            return $false
        }
        finally {
            if ($mail) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($mail) | Out-Null }
            if ($outlook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null }
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }
}
# --- ENDE: Erforderliche Funktionen ---


# --- START: Ausführungsdaten ---
$UserList = @($userListString)
$RDPFilePaths = @($rdpListString)
$OutputPath = "$resolvedOutputPath"
$SmtpServer = $smtpServerString
$SendAction = $sendActionString
# --- ENDE: Ausführungsdaten ---

Write-Host "Starte E-Mail-Versand-Skript..."
Write-Host "Betrifft ($($UserList.Count)) Benutzer."
Write-Host "Versandmethode: " -NoNewline
if ($SmtpServer) {
    Write-Host "SMTP ($SmtpServer)" -ForegroundColor Yellow
} else {
    Write-Host "Outlook COM" -ForegroundColor Cyan
}

foreach ($userName in $UserList) {
    $msgFileName = "$($userName).msg"
    $msgFullPath = if (-not $SmtpServer) { Join-Path -Path $OutputPath -ChildPath $msgFileName } else { $null }

    Write-Host "--- Verarbeite $userName ---"
    Start-EmailWorkflow -UserName $userName -RDPFilePaths $RDPFilePaths -MSGSavePath $msgFullPath -Send:$SendAction -SmtpServer $SmtpServer -Credential $null
}

Write-Host "E-Mail-Versand-Skript beendet."
Read-Host "Drücken Sie ENTER, um das Fenster zu schließen."
"@
    
    try {
        $scriptFileName = "sendMails_$($GlobalTimestamp).ps1"
        $scriptFullPath = Join-Path -Path $OutputPath -ChildPath $scriptFileName
        Set-Content -Path $scriptFullPath -Value $scriptContent -Encoding UTF8 -ErrorAction Stop
        
        Write-Host "-------------------------------------------------------------" -ForegroundColor Cyan
        Write-Host "[OK] 'sendMails.ps1' wurde generiert!" -ForegroundColor Green
        Write-Warning "HINWEIS: Wenn der direkte Outlook-Versand (als Admin) fehlschlägt (CO_E_SERVER_EXEC_FAILURE):"
        Write-Warning "1. Öffnen Sie eine NEUE, NORMALE (Nicht-Admin) PowerShell-Konsole."
        Write-Warning "2. Führen Sie das folgende Skript in dieser neuen Konsole aus:"
        Write-Host "   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force" -ForegroundColor White
        Write-Host "   '&' '$scriptFullPath'" -ForegroundColor White
        Write-Host "-------------------------------------------------------------"
    } catch {
        $msg = "Fehler beim Generieren des 'sendMails.ps1'-Skripts: $($_.Exception.Message)"
        Write-Warning $msg
        $GlobalErrorLog.Add($msg) | Out-Null
    }
}


#==============================================================================
# LOGIK-BLOCKS
#==============================================================================

# --- LOGIK-BLOCK 1 & 2: BERECHTIGUNGEN (N:M) ---
Function Start-RightsWorkflow {
    param(
        [bool]$SetRights,
        [bool]$RemoveRights,
        [string]$UserListPath,
        [string]$ClientListPath,
        [string]$UserColumn,
        [string]$ClientColumn,
        [switch]$SendEmail,
        [string]$SmtpServer,
        [System.Management.Automation.PSCredential]$Credential,
        [string]$OutputPath
    )
    
    $Users = Load-CsvData -CsvPath $UserListPath
    $Clients = Load-CsvData -CsvPath $ClientListPath
    if ($null -eq $Users -or $null -eq $Clients) { return }
    
    if (-not (Test-ADModule)) { return }
    $ADContext = Get-DynamicADDomain
    
    $CurrentAction = if ($SetRights) { 'Add' } else { 'Remove' }
    $ActionVerb = if ($SetRights) { "Hinzufügen" } else { "Entfernen" }
    
    $userCount = ($Users | Measure-Object).Count
    $clientCount = ($Clients | Measure-Object).Count
    
    Write-Host "Starte Modus: Remotedesktop-Berechtigungen ($ActionVerb) (Alle User -> Alle Clients)" -ForegroundColor Cyan
    Write-Host "Betrifft $userCount Benutzer auf $clientCount Clients."

    # --- PHASE 1: PLANUNG & VORAB-TESTS ---
    Write-Host "PHASE 1: Plane Aktionen und teste Erreichbarkeit (Ping)..." -ForegroundColor Yellow
    $plannedActions = @()
    $offlineClients = @()

    foreach ($client in $Clients) {
        $clientName = $client.$($ClientColumn)
        if ([string]::IsNullOrWhiteSpace($clientName)) { Write-Warning "Eintrag in Client-Liste übersprungen (Name ist leer)."; continue }
        
        Write-Verbose "Teste Erreichbarkeit von $clientName..."
        if (-not (Test-Connection -ComputerName $clientName -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
            $msg = "Client $clientName ist offline (Ping fehlgeschlagen). Alle Aktionen für diesen Client werden übersprungen."
            Write-Warning $msg
            $GlobalErrorLog.Add($msg) | Out-Null
            $offlineClients += $clientName
            continue
        }
        
        Write-Verbose "Client $clientName ist online. Plane Aktionen..."
        foreach ($user in $Users) {
            $userName = $user.$($UserColumn)
            if ([string]::IsNullOrWhiteSpace($userName)) { Write-Warning "Eintrag in Benutzer-Liste übersprungen (Name ist leer)."; continue }
            
            $plannedActions += [PSCustomObject]@{
                Client     = $clientName
                User       = $userName
                Action     = $CurrentAction
            }
        }
    }
    
    # --- PHASE 2: ANZEIGE & BESTÄTIGUNG ---
    Write-Host "-------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host "PHASE 2: Planungsübersicht und Bestätigung" -ForegroundColor Cyan
    $offlineCount = ($offlineClients | Measure-Object).Count
    if ($offlineCount -gt 0) {
        Write-Host "Die folgenden $offlineCount Clients sind OFFLINE und werden übersprungen:" -ForegroundColor Yellow
        $offlineClients | ForEach-Object { Write-Host " - $_" }
    }
    
    $plannedCount = ($plannedActions | Measure-Object).Count
    if ($plannedCount -eq 0) {
        Write-Host "Keine Aktionen für Online-Clients geplant." -ForegroundColor Green
        Write-Host "Modus beendet."
        Write-ErrorLog -ErrorOutputPath $GlobalLogDir 
        return
    }

    Write-Host "Es sind $plannedCount Aktionen für Online-Clients geplant:" -ForegroundColor Green
    $plannedActions | Format-Table -AutoSize
    Write-Host "-------------------------------------------------------------"
    
    # --- FIX V3.0: Robuste Bestätigungslogik ---
    $doExecute = $false
    # $pscmdlet.ShouldProcess prüft auf -WhatIf
    if ($pscmdlet.ShouldProcess("die $plannedCount oben gelisteten Aktionen", "Ausführen")) {
        
        # Prüft, ob -Confirm:$false explizit genutzt wurde
        if ($PSBoundParameters.ContainsKey('Confirm') -and (-not $Confirm)) {
            Write-Verbose "Bestätigung erzwungen durch -Confirm:`$false."
            $doExecute = $true
        } else {
            # Standardfall: -WhatIf ist NICHT gesetzt, -Confirm:$false ist NICHT gesetzt.
            # IMMER fragen.
            $confirmation = Read-Host -Prompt "Möchten Sie diese $plannedCount Aktionen jetzt ausführen? (J/N)"
            if ($confirmation.ToUpper() -eq 'J') {
                $doExecute = $true
            }
        }
    }
    # $doExecute ist $false, wenn der Benutzer 'N' gesagt hat oder -WhatIf verwendet hat.
    
    if ($doExecute) {
        # --- PHASE 3: AUSFÜHRUNG ---
        Write-Host "-------------------------------------------------------------" -ForegroundColor Cyan
        Write-Host "PHASE 3: Führe Aktionen aus..." -ForegroundColor Cyan
        
        $Report = @()
        foreach ($action in $plannedActions) {
            $status = Invoke-RemoteGroupMembership -ComputerName $action.Client -UserName $action.User -Domain $ADContext.NetBIOS -Action $action.Action
            
            switch ($status) {
                'Success' {
                    if ($action.Action -eq 'Add') {
                        Write-Host " ERFOLG: $($action.User) zu $($action.Client) HINZUGEFÜGT." -ForegroundColor Green
                    } else {
                        Write-Host " ERFOLG: $($action.User) von $($action.Client) ENTFERNT." -ForegroundColor Yellow
                    }
                }
                'AlreadyExists' { Write-Host " STATUS: $($action.User) ist bereits Mitglied auf $($action.Client)." -ForegroundColor Gray }
                'NotMember' { Write-Host " STATUS: $($action.User) ist nicht Mitglied auf $($action.Client)." -ForegroundColor Gray }
                'VerificationFailed' { Write-Warning " FEHLER: Aktion für $($action.User) auf $($action.Client) ausgeführt, aber Verifizierung schlug fehl." }
                'Failed' { # Fehler wurde bereits in der Funktion geloggt
                }
            }
            
            $Report += [PSCustomObject]@{
                Client     = $action.Client
                User       = $action.User
                Action     = $action.Action
                Status     = $status
                Timestamp  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
        
        Write-Host "Modus ($ActionVerb) abgeschlossen." -ForegroundColor Cyan
        $Report | Format-Table
        
        Write-Log -ReportData $Report
        Write-ErrorLog -ErrorOutputPath $GlobalLogDir

        if ($SendEmail) {
            Write-Host "-------------------------------------------------------------" -ForegroundColor Cyan
            Write-Host "Kombinierter Modus: Starte E-Mail-Workflow..." -ForegroundColor Cyan
            
            $successfulReport = $Report | Where-Object { $_.Status -in ('Success', 'AlreadyExists') }
            $successCount = ($successfulReport | Measure-Object).Count
            
            if ($successCount -gt 0) {
                $successUsers = $successfulReport | Select-Object -ExpandProperty User -Unique
                $successClients = $successfulReport | Select-Object -ExpandProperty Client -Unique
                
                Start-FileWorkflow -UserList $successUsers -ClientList $successClients -OutputPath $OutputPath -UserColumn "User" -ClientColumn "Client" -SmtpServer $SmtpServer -Credential $Credential -IsCombinedRun $true -SaveAsMsgOnly:$SaveAsMsgOnly
            } else {
                $msg = "E-Mail-Workflow übersprungen, da keine Rechte erfolgreich zugewiesen wurden."
                Write-Warning $msg
                $GlobalErrorLog.Add($msg) | Out-Null
            }
        }
    }
    else {
        Write-Warning "Aktion vom Benutzer abgebrochen oder -WhatIf verwendet."
    }
}


# --- LOGIK-BLOCK 3 & 4: RDP/MSG-DATEIEN ---
Function Start-FileWorkflow {
    param(
        # Modus 3
        [string]$UserListPath,
        [string]$ClientListPath,
        # Modus 4
        [string]$InputLogPath,
        [Switch]$GenerateFromLog, 
        # Kombi-Modus (1+3)
        [System.Array]$UserList,
        [System.Array]$ClientList,
        [bool]$IsCombinedRun,
        # Gemeinsame
        [string]$OutputPath,
        [string]$UserColumn,
        [string]$ClientColumn,
        [switch]$SaveAsMsgOnly,
        [string]$SmtpServer,
        [System.Management.Automation.PSCredential]$Credential
    )
    
    $Users = $null
    $Clients = $null

    if ($IsCombinedRun) {
        $Users = $UserList
        $Clients = $ClientList
        $UserColumn = "User" 
        $ClientColumn = "Client"
    }
    elseif ($GenerateFromLog.IsPresent) { 
        $LogData = Load-CsvData -CsvPath $InputLogPath -ExpectedHeader "Status"
        if ($null -eq $LogData) { return }
        $successfulReport = $LogData | Where-Object { $_.Status -in ('Success', 'AlreadyExists') }
        $Users = $successfulReport | Select-Object -ExpandProperty User -Unique
        $Clients = $successfulReport | Select-Object -ExpandProperty Client -Unique
        $UserColumn = "User" 
        $ClientColumn = "Client"
    }
    else {
        $Users = Load-CsvData -CsvPath $UserListPath
        $Clients = Load-CsvData -CsvPath $ClientListPath
        if ($null -eq $Users -or $null -eq $Clients) { return }
    }
    
    if (-not (Test-ADModule)) { return }

    $userCount = ($Users | Measure-Object).Count
    $clientCount = ($Clients | Measure-Object).Count
    
    if ($userCount -eq 0 -or $clientCount -eq 0) {
        $msg = "Keine gültigen Benutzer ($userCount) oder Clients ($clientCount) für den E-Mail-Workflow gefunden. Abbruch."
        Write-Error $msg
        $GlobalErrorLog.Add($msg) | Out-Null
        return
    }
    
    Write-Host "Starte Modus: RDP/E-Mail-Dateien erstellen (N:M Logik)" -ForegroundColor Cyan
    
    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        $OutputPath = Join-Path -Path $PSScriptRoot -ChildPath "RDP_Ausgabe_$($GlobalTimestamp)"
    }
    
    if (-not (Test-Path $OutputPath)) {
        try {
            Write-Verbose "Erstelle Ausgabeordner: $OutputPath"
            New-Item -Path $OutputPath -ItemType Directory -ErrorAction Stop | Out-Null
        } catch {
            $msg = "Ausgabeordner '$OutputPath' konnte nicht erstellt werden. Bitte manuell anlegen. Abbruch."
            Write-Error $msg
            $GlobalErrorLog.Add($msg) | Out-Null
            return
        }
    }
    
    # --- PHASE 1: PLANUNG & VORAB-TESTS ---
    Write-Host "PHASE 1: Plane Aktionen..." -ForegroundColor Yellow
    $onlineClients = @()
    $offlineClients = @()
    $rdpFilePaths = @()
    
    $clientListObjects = if ($IsCombinedRun -or $GenerateFromLog) { $Clients } else { $Clients }
    $userListObjects = if ($IsCombinedRun -or $GenerateFromLog) { $Users } else { $Users }
    
    foreach ($client in $clientListObjects) {
        $clientName = if ($IsCombinedRun -or $GenerateFromLog) { $client } else { $client.$($ClientColumn) }
        if ([string]::IsNullOrWhiteSpace($clientName)) { continue }
        
        Write-Verbose "Teste Erreichbarkeit von $clientName..."
        if (-not (Test-Connection -ComputerName $clientName -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
            $msg = "Client $clientName ist offline. Es wird keine RDP-Datei dafür erstellt."
            Write-Warning $msg
            $GlobalErrorLog.Add($msg) | Out-Null
            $offlineClients += $clientName
            continue
        }
        $onlineClients += $clientName
    }
    
    $onlineCount = ($onlineClients | Measure-Object).Count
    if ($onlineCount -eq 0) {
        $msg = "Keine Clients online. Abbruch."
        Write-Error $msg
        $GlobalErrorLog.Add($msg) | Out-Null
        Write-ErrorLog -ErrorOutputPath $OutputPath
        return
    }
    
    # --- PHASE 2: ANZEIGE & BESTÄTIGUNG ---
    Write-Host "-------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host "PHASE 2: Planungsübersicht und Bestätigung" -ForegroundColor Cyan
    $offlineCount = ($offlineClients | Measure-Object).Count
    if ($offlineCount -gt 0) {
        Write-Host "Die folgenden $offlineCount Clients sind OFFLINE und werden übersprungen:" -ForegroundColor Yellow
        $offlineClients | ForEach-Object { Write-Host " - $_" }
    }
    
    $totalRDPs = $onlineCount
    $totalEmails = $userCount
    $totalActions = $totalRDPs + $totalEmails
    
    Write-Host "Geplante Aktionen:" -ForegroundColor Green
    Write-Host " - $($totalRDPs) RDP-Dateien werden erstellt (eine pro Online-Client)."
    Write-Host " - $($totalEmails) E-Mails werden erstellt (eine pro Benutzer)."
    
    $sendAction = $false
    if (-not $SaveAsMsgOnly) {
        $sendAction = $true
        if (-not [string]::IsNullOrWhiteSpace($SmtpServer)) {
            Write-Host " - Die $($totalEmails) E-Mails werden DIREKT VIA SMTP ($SmtpServer) VERSENDET." -ForegroundColor Red
        } else {
            Write-Host " - Die $($totalEmails) E-Mails werden DIREKT VIA OUTLOOK VERSENDET (und als .msg gespeichert)." -ForegroundColor Red
        }
    } else {
        Write-Host " - Die $($totalEmails) E-Mails werden nur als .MSG im Ausgabeordner GESPEICHERT." -ForegroundColor Yellow
    }
    Write-Host "-------------------------------------------------------------"

    # --- FIX V2.8: Robuste Bestätigungslogik ---
    $doExecute = $false
    if ($IsCombinedRun) {
        # Bereits in Start-RightsWorkflow bestätigt
        $doExecute = $true
    }
    elseif ($pscmdlet.ShouldProcess("die $totalActions oben gelisteten Datei-Workflows", "Ausführen")) {
        if ($PSBoundParameters.ContainsKey('Confirm') -and -not $Confirm) {
            Write-Verbose "Bestätigung erzwungen durch -Confirm:`$false."
            $doExecute = $true
        } else {
            $confirmation = Read-Host -Prompt "Möchten Sie diese $totalActions Aktionen jetzt ausführen? (J/N)"
            if ($confirmation.ToUpper() -eq 'J') {
                $doExecute = $true
            }
        }
    }

    if ($doExecute) {
        # --- PHASE 3: AUSFÜHRUNG ---
        Write-Host "-------------------------------------------------------------" -ForegroundColor Cyan
        Write-Host "PHASE 3: Führe Workflows aus..." -ForegroundColor Cyan
        $Report = @()
        
        # 1. RDP-Dateien erstellen
        Write-Host "Erstelle RDP-Dateien..."
        foreach ($clientName in $onlineClients) {
            $rdpFileName = "$($clientName).rdp"
            $rdpFullPath = Join-Path -Path $OutputPath -ChildPath $rdpFileName
            if (Create-RDPFile -ComputerName $clientName -FilePath $rdpFullPath) {
                $rdpFilePaths += $rdpFullPath 
            }
        }
        
        $rdpCount = ($rdpFilePaths | Measure-Object).Count
        if ($rdpCount -eq 0) {
            $msg = "Konnte keine RDP-Dateien erstellen. E-Mail-Versand wird übersprungen."
            Write-Error $msg
            $GlobalErrorLog.Add($msg) | Out-Null
            Write-ErrorLog -ErrorOutputPath $OutputPath
            return
        }

        # 2. E-Mails erstellen/senden
        Write-Host "Erstelle/Sende E-Mails..."
        foreach ($user in $userListObjects) {
            $userName = if ($IsCombinedRun -or $GenerateFromLog) { $user } else { $user.$($UserColumn) }
            if ([string]::IsNullOrWhiteSpace($userName)) { continue }
            
            $msgFullPath = $null
            if ([string]::IsNullOrWhiteSpace($SmtpServer)) {
                $msgFileName = "$($userName).msg"
                $msgFullPath = Join-Path -Path $OutputPath -ChildPath $msgFileName
            }
            
            $emailSuccess = Start-EmailWorkflow -UserName $userName -RDPFilePaths $rdpFilePaths -MSGSavePath $msgFullPath -Send:$sendAction -SmtpServer $SmtpServer -Credential $Credential
            
            $Report += [PSCustomObject]@{
                User       = $userName
                Action     = if ($sendAction) { "Email_Send" } else { "Email_Save_MSG" }
                Status     = if ($emailSuccess) { "Success" } else { "Failed" }
                Timestamp  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
        
        # 3. 'sendMails.ps1' Skript generieren
        Generate-SendMailsScript -UserList $userListObjects -RDPFilePaths $rdpFilePaths -OutputPath $OutputPath -SmtpServer $SmtpServer -SaveAsMsgOnly:$SaveAsMsgOnly
        
        Write-Host "Modus (Datei-Workflow) abgeschlossen." -ForegroundColor Cyan
        if (($Report | Measure-Object).Count -gt 0) {
            $Report | Format-Table
            Write-Log -ReportData $Report
        }
        
        Write-ErrorLog -ErrorOutputPath $OutputPath
        Write-Host "Alle Ausgaben finden Sie in: $OutputPath"
    }
    else {
        Write-Warning "Aktion vom Benutzer abgebrochen oder -WhatIf verwendet."
    }
}

#==============================================================================
# SCRIPT-STARTPUNKT
#==============================================================================

# --- MODUS 1: PARAMETER-MODUS ---
if ($PSCmdlet.ParameterSetName -ne 'Interactive') {
    
    Write-Host "Skript im Parameter-Modus gestartet (Modus: $($PSCmdlet.ParameterSetName))." -ForegroundColor Yellow
    
    if ($SetRDPRights -or $RemoveRDPRights) {
        Start-RightsWorkflow -SetRights $SetRDPRights -RemoveRights $RemoveRDPRights -UserListPath $UserListPath -ClientListPath $ClientListPath -UserColumn $UserColumn -ClientColumn $ClientColumn -SendEmail:$SendEmail -SmtpServer $SmtpServer -Credential $Credential -OutputPath $OutputPath
    }
    elseif ($GenerateRDPFiles) {
        Start-FileWorkflow -GenerateRDPFiles -UserListPath $UserListPath -ClientListPath $ClientListPath -OutputPath $OutputPath -UserColumn $UserColumn -ClientColumn $ClientColumn -SaveAsMsgOnly:$SaveAsMsgOnly -SmtpServer $SmtpServer -Credential $Credential
    }
    elseif ($GenerateFromLog) {
        Start-FileWorkflow -GenerateFromLog:$true -InputLogPath $InputLogPath -OutputPath $OutputPath -SaveAsMsgOnly:$SaveAsMsgOnly -SmtpServer $SmtpServer -Credential $Credential
    }
}
# --- MODUS 2: INTERAKTIVES MENÜ ---
else {
    
    while ($true) {
        $GlobalErrorLog.Clear()
        $choice = Show-MainMenu
        
        switch ($choice) {
            '1' { # Rechte HINZUFÜGEN
                $inputParams = Get-InteractiveInput -Mode 'SetRDPRights'
                Start-RightsWorkflow -SetRights $true -UserListPath $inputParams.UserListPath -ClientListPath $inputParams.ClientListPath -UserColumn $inputParams.UserColumn -ClientColumn $inputParams.ClientColumn -SendEmail:$inputParams.SendEmail -SmtpServer $inputParams.SmtpServer -Credential $inputParams.Credential -OutputPath $inputParams.OutputPath
            }
            '2' { # Rechte ENTFERNEN
                $inputParams = Get-InteractiveInput -Mode 'RemoveRDPRights'
                Start-RightsWorkflow -RemoveRights $true -UserListPath $inputParams.UserListPath -ClientListPath $inputParams.ClientListPath -UserColumn $inputParams.UserColumn -ClientColumn $inputParams.ClientColumn -SendEmail:$inputParams.SendEmail -SmtpServer $inputParams.SmtpServer -Credential $inputParams.Credential -OutputPath $inputParams.OutputPath
            }
            '3' { # Dateien ERSTELLEN (aus CSV)
                $inputParams = Get-InteractiveInput -Mode 'GenerateRDPFiles'
                Start-FileWorkflow -GenerateRDPFiles -UserListPath $inputParams.UserListPath -ClientListPath $inputParams.ClientListPath -OutputPath $inputParams.OutputPath -UserColumn $inputParams.UserColumn -ClientColumn $inputParams.ClientColumn -SaveAsMsgOnly:$inputParams.SaveAsMsgOnly -SmtpServer $inputParams.SmtpServer -Credential $inputParams.Credential
            }
            '4' { # Dateien ERSTELLEN (aus Log)
                $inputParams = Get-InteractiveInput -Mode 'GenerateFromLog'
                Start-FileWorkflow -GenerateFromLog:$true -InputLogPath $inputParams.InputLogPath -OutputPath $inputParams.OutputPath -SaveAsMsgOnly:$inputParams.SaveAsMsgOnly -SmtpServer $inputParams.SmtpServer -Credential $inputParams.Credential
            }
            'Q' { Write-Host "Beendet." -ForegroundColor Gray; return }
            default { Write-Warning "Ungültige Auswahl." }
        }
        
        Write-Host "-------------------------------------------------------------" -ForegroundColor Gray
        Read-Host -Prompt "Drücken Sie ENTER, um zum Hauptmenü zurückzukehren..."
    }
}

Write-Host "Skript-Ausführung beendet."
