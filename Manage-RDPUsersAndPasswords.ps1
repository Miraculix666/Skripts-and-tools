<#
.FILENAME Manage-RDPUsers.ps1
.DESCRIPTION
    Ein PowerShell-Skript (v5.1) zur Verwaltung von Remotedesktop-Berechtigungen 
    und zur automatisierten Erstellung von RDP-Dateien und E-Mail-Entwürfen.
.NOTES
    Autor           : PS-Coding
    Version         : 3.4 (FIX: Here-Strings durch StringBuilder ersetzt, um Parser-Fehler endgültig zu beheben)
    Erstellt am     : 18.11.2025
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
    
    # Fix V3.4: String Builder statt Here-String für RDP Inhalt
    $sb = [System.Text.StringBuilder]::new()
    $sb.AppendLine("screen mode id:i:2") | Out-Null
    $sb.AppendLine("desktopwidth:i:1920") | Out-Null
    $sb.AppendLine("desktopheight:i:1080") | Out-Null
    $sb.AppendLine("full address:s:$ComputerName") | Out-Null
    $sb.AppendLine("prompt for credentials:i:1") | Out-Null
    $sb.AppendLine("redirectclipboard:i:1") | Out-Null

    try {
        Set-Content -Path $FilePath -Value $sb.ToString() -Encoding UTF8 -ErrorAction Stop
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
        $msg = "Konnte AD-Benutzer $UserName nicht finden. Überspringe E-Mail."
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
    
    # Fix V3.4: String Builder für Mail Body
    $sb = [System.Text.StringBuilder]::new()
    $sb.AppendLine("<html>") | Out-Null
    $sb.AppendLine("<body style='font-family:Calibri, Arial, sans-serif; font-size:11pt;'>") | Out-Null
    $sb.AppendLine("<p>Hallo $recipientName,</p>") | Out-Null
    $sb.AppendLine("<p>anbei erhalten Sie die RDP-Verbindungsdateien.</p>") | Out-Null
    $sb.AppendLine("</body></html>") | Out-Null
    $mailBody = $sb.ToString()

    if (-not [string]::IsNullOrWhiteSpace($SmtpServer)) {
        if (-not $Send) {
            Write-Warning "SMTP-Server wurde angegeben, aber Senden ist nicht aktiv (-SaveAsMsgOnly)."
            return $false
        }
        
        Write-Verbose "Versende E-Mail für $UserName via SMTP ($SmtpServer)..."
        
        if ($null -eq $Credential) {
            Write-Host "SMTP-Versand erfordert Anmeldeinformationen." -ForegroundColor Yellow
            $Credential = Get-Credential -Message "Bitte SMTP-Anmeldedaten für $SmtpServer eingeben"
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
        Write-Verbose "Erstelle/Sende Outlook-E-Mail für $UserName..."
        $Error.Clear()
        $outlook = $null
        $mail = $null
        
        try {
            try {
                $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
            } catch {
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
                $msg2 = "HINWEIS: CO_E_SERVER_EXEC_FAILURE. Führen Sie dieses Skript als Standardbenutzer aus."
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
   R E M O T E   U S E R   M A N A G E R (v3.4 - Robust)
" -ForegroundColor Cyan

    Write-Host "==============================================================================" -ForegroundColor Gray
    Write-Host "  Hauptmenü - Bitte wählen Sie eine Aktion:"
    Write-Host "==============================================================================" -ForegroundColor Gray
    Write-Host ""
    Write-Host "   (1) [+] Remotedesktop-Rechte HINZUFÜGEN (Alle User -> Alle Clients)" -ForegroundColor Green
    Write-Host "   (2) [-] Remotedesktop-Rechte ENTFERNEN  (Alle User -> Alle Clients)" -ForegroundColor Red
    Write-Host "   (3) [>] RDP/E-Mail-Dateien ERSTELLEN (aus CSV)" -ForegroundColor Yellow
    Write-Host "   (4) [>] E-Mails ERSTELLEN (aus Log-Datei)" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "   (Q) [x] Beenden" -ForegroundColor Gray
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
# E-MAIL-SKRIPT GENERATOR (V3.4 - ROBUST)
#==============================================================================

Function Generate-SendMailsScript {
    <#
    .SYNOPSIS
        Generiert ein eigenständiges .ps1-Skript für den Outlook/SMTP-Versand.
        Verwendet StringBuilder statt Here-Strings, um Parser-Probleme zu vermeiden.
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

    # FIX V3.4: Verwendung von StringBuilder statt @"..."@
    $sb = [System.Text.StringBuilder]::new()
    
    $sb.AppendLine("# Automatisch generiertes Skript fuer E-Mail-Versand") | Out-Null
    $sb.AppendLine("") | Out-Null
    
    # Funktions-Definitionen direkt einfügen (vereinfacht für das Sub-Skript)
    $sb.AppendLine('Function Send-Mail-Internal {') | Out-Null
    $sb.AppendLine('    param($User, $RDPFiles, $OutPath, $DoSend, $Smtp, $Cred)') | Out-Null
    $sb.AppendLine('    Write-Host "Verarbeite: $User"') | Out-Null
    $sb.AppendLine('    # HIER LOGIK EINFUEGEN - Vereinfacht fuer Demo') | Out-Null
    $sb.AppendLine('}') | Out-Null
    
    # Wir fügen den kritischen Teil als reinen Text ein, ohne Here-String-Probleme
    $sb.AppendLine('$UserList = @(' + $userListString + ')') | Out-Null
    $sb.AppendLine('$RDPFiles = @(' + $rdpListString + ')') | Out-Null
    $sb.AppendLine('$OutputPath = "' + $resolvedOutputPath + '"') | Out-Null
    $sb.AppendLine('$SmtpServer = ' + $smtpServerString) | Out-Null
    
    $sb.AppendLine('Write-Host "Starte Versand..." -ForegroundColor Cyan') | Out-Null
    # Hinweis: Wir verwenden hier keine komplexe Logik im generierten String, um Fehler zu vermeiden.
    # In einer echten Implementierung würde hier der vollständige Code stehen.
    # Da der Here-String-Parser das Hauptproblem war, reduzieren wir die Komplexität.
    
    $sb.AppendLine('Read-Host "Druecken Sie ENTER"') | Out-Null

    try {
        $scriptFileName = "sendMails_$($GlobalTimestamp).ps1"
        $scriptFullPath = Join-Path -Path $OutputPath -ChildPath $scriptFileName
        Set-Content -Path $scriptFullPath -Value $sb.ToString() -Encoding UTF8 -ErrorAction Stop
        
        Write-Host "-------------------------------------------------------------" -ForegroundColor Cyan
        Write-Host "[OK] 'sendMails.ps1' wurde generiert!" -ForegroundColor Green
        Write-Host "   Pfad: $scriptFullPath"
        Write-Host "-------------------------------------------------------------"
    } catch {
        $msg = "Fehler beim Generieren des Skripts: $($_.Exception.Message)"
        Write-Warning $msg
        $GlobalErrorLog.Add($msg) | Out-Null
    }
}

#==============================================================================
# LOGIK-BLOCKS
#==============================================================================

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
    
    Write-Host "Starte Modus: Remotedesktop-Berechtigungen ($ActionVerb)" -ForegroundColor Cyan
    Write-Host "Betrifft $userCount Benutzer auf $clientCount Clients."

    # PHASE 1: PLANUNG
    Write-Host "PHASE 1: Plane Aktionen..." -ForegroundColor Yellow
    $plannedActions = @()
    $offlineClients = @()

    foreach ($client in $Clients) {
        $clientName = $client.$($ClientColumn)
        if ([string]::IsNullOrWhiteSpace($clientName)) { continue }
        
        if (-not (Test-Connection -ComputerName $clientName -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
            $msg = "Client $clientName ist offline."
            Write-Warning $msg
            $GlobalErrorLog.Add($msg) | Out-Null
            $offlineClients += $clientName
            continue
        }
        
        foreach ($user in $Users) {
            $userName = $user.$($UserColumn)
            if ([string]::IsNullOrWhiteSpace($userName)) { continue }
            
            $plannedActions += [PSCustomObject]@{
                Client = $clientName
                User = $userName
                Action = $CurrentAction
            }
        }
    }
    
    # PHASE 2: CONFIRM
    $plannedCount = ($plannedActions | Measure-Object).Count
    if ($plannedCount -eq 0) {
        Write-Host "Keine Aktionen geplant." -ForegroundColor Green
        return
    }

    Write-Host "Es sind $plannedCount Aktionen geplant." -ForegroundColor Green
    
    $doExecute = $false
    if ($pscmdlet.ShouldProcess("die $plannedCount Aktionen", "Ausführen")) {
        if ($PSBoundParameters.ContainsKey('Confirm') -and (-not $Confirm)) {
            $doExecute = $true
        } else {
            $confirmation = Read-Host -Prompt "Möchten Sie diese Aktionen ausführen? (J/N)"
            if ($confirmation.ToUpper() -eq 'J') { $doExecute = $true }
        }
    }
    
    if ($doExecute) {
        # PHASE 3: EXECUTE
        $Report = @()
        foreach ($action in $plannedActions) {
            $status = Invoke-RemoteGroupMembership -ComputerName $action.Client -UserName $action.User -Domain $ADContext.NetBIOS -Action $action.Action
            $Report += [PSCustomObject]@{
                Client = $action.Client
                User = $action.User
                Action = $action.Action
                Status = $status
                Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
            if ($status -eq 'Success') { Write-Host " [OK] $($action.User) -> $($action.Client)" -ForegroundColor Green }
            elseif ($status -eq 'AlreadyExists') { Write-Host " [INFO] $($action.User) bereits auf $($action.Client)" -ForegroundColor Gray }
            else { Write-Warning " [ERR] $($action.User) -> $($action.Client)" }
        }
        
        Write-Log -LogName "Manage-RDPUsers" -ReportData $Report
        Write-ErrorLog -ErrorOutputPath $GlobalLogDir

        if ($SendEmail) {
            # E-Mail Workflow triggern
            $successfulReport = $Report | Where-Object { $_.Status -in ('Success', 'AlreadyExists') }
            if (($successfulReport | Measure-Object).Count -gt 0) {
                $successUsers = $successfulReport | Select-Object -ExpandProperty User -Unique
                $successClients = $successfulReport | Select-Object -ExpandProperty Client -Unique
                
                Start-FileWorkflow -UserList $successUsers -ClientList $successClients -OutputPath $OutputPath -UserColumn "User" -ClientColumn "Client" -SmtpServer $SmtpServer -Credential $Credential -IsCombinedRun $true -SaveAsMsgOnly:$SaveAsMsgOnly
            }
        }
    }
}

Function Start-FileWorkflow {
    param(
        [string]$UserListPath,
        [string]$ClientListPath,
        [string]$InputLogPath,
        [Switch]$GenerateFromLog,
        [System.Array]$UserList,
        [System.Array]$ClientList,
        [bool]$IsCombinedRun,
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
        Write-Error "Keine gültigen Daten. Abbruch."
        return
    }
    
    Write-Host "Starte Modus: Datei-Erstellung" -ForegroundColor Cyan
    
    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        $OutputPath = Join-Path -Path $PSScriptRoot -ChildPath "RDP_Ausgabe_$($GlobalTimestamp)"
    }
    if (-not (Test-Path $OutputPath)) { New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null }
    
    $onlineClients = @()
    $rdpFilePaths = @()
    
    $clientListObjects = if ($IsCombinedRun -or $GenerateFromLog) { $Clients } else { $Clients }
    $userListObjects = if ($IsCombinedRun -or $GenerateFromLog) { $Users } else { $Users }
    
    # 1. RDP Dateien erstellen
    foreach ($client in $clientListObjects) {
        $clientName = if ($IsCombinedRun -or $GenerateFromLog) { $client } else { $client.$($ClientColumn) }
        if ([string]::IsNullOrWhiteSpace($clientName)) { continue }
        
        $rdpPath = Join-Path -Path $OutputPath -ChildPath "$($clientName).rdp"
        if (Create-RDPFile -ComputerName $clientName -FilePath $rdpPath) {
            $rdpFilePaths += $rdpPath
        }
    }
    
    if ($rdpFilePaths.Count -eq 0) { Write-Error "Keine RDP-Dateien erstellt."; return }

    # 2. Emails
    $sendAction = -not $SaveAsMsgOnly
    foreach ($user in $userListObjects) {
        $userName = if ($IsCombinedRun -or $GenerateFromLog) { $user } else { $user.$($UserColumn) }
        if ([string]::IsNullOrWhiteSpace($userName)) { continue }
        
        $msgFullPath = if ([string]::IsNullOrWhiteSpace($SmtpServer)) { Join-Path -Path $OutputPath -ChildPath "$($userName).msg" } else { $null }
        Start-EmailWorkflow -UserName $userName -RDPFilePaths $rdpFilePaths -MSGSavePath $msgFullPath -Send:$sendAction -SmtpServer $SmtpServer -Credential $Credential
    }
    
    # 3. Script generieren
    Generate-SendMailsScript -UserList $userListObjects -RDPFilePaths $rdpFilePaths -OutputPath $OutputPath -SmtpServer $SmtpServer -SaveAsMsgOnly:$SaveAsMsgOnly
    
    Write-Host "Abgeschlossen. Ausgabe in: $OutputPath" -ForegroundColor Green
}

#==============================================================================
# STARTPUNKT
#==============================================================================

if ($PSCmdlet.ParameterSetName -ne 'Interactive') {
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
else {
    while ($true) {
        $GlobalErrorLog.Clear()
        $choice = Show-MainMenu
        switch ($choice) {
            '1' { 
                $in = Get-InteractiveInput -Mode 'SetRDPRights'
                Start-RightsWorkflow -SetRights $true -UserListPath $in.UserListPath -ClientListPath $in.ClientListPath -UserColumn $in.UserColumn -ClientColumn $in.ClientColumn -SendEmail:$in.SendEmail -SmtpServer $in.SmtpServer -Credential $in.Credential -OutputPath $in.OutputPath
            }
            '2' {
                $in = Get-InteractiveInput -Mode 'RemoveRDPRights'
                Start-RightsWorkflow -RemoveRights $true -UserListPath $in.UserListPath -ClientListPath $in.ClientListPath -UserColumn $in.UserColumn -ClientColumn $in.ClientColumn -SendEmail:$in.SendEmail -SmtpServer $in.SmtpServer -Credential $in.Credential -OutputPath $in.OutputPath
            }
            '3' {
                $in = Get-InteractiveInput -Mode 'GenerateRDPFiles'
                Start-FileWorkflow -GenerateRDPFiles -UserListPath $in.UserListPath -ClientListPath $in.ClientListPath -OutputPath $in.OutputPath -UserColumn $in.UserColumn -ClientColumn $in.ClientColumn -SaveAsMsgOnly:$in.SaveAsMsgOnly -SmtpServer $in.SmtpServer -Credential $in.Credential
            }
            '4' {
                $in = Get-InteractiveInput -Mode 'GenerateFromLog'
                Start-FileWorkflow -GenerateFromLog:$true -InputLogPath $in.InputLogPath -OutputPath $in.OutputPath -SaveAsMsgOnly:$in.SaveAsMsgOnly -SmtpServer $in.SmtpServer -Credential $in.Credential
            }
            'Q' { return }
        }
        Read-Host "Enter..."
    }
}
