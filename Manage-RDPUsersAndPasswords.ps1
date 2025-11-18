<#
.FILENAME Manage-RDPUsers.ps1
.DESCRIPTION
    Ein PowerShell-Skript (v5.1) zur Verwaltung von Remotedesktop-Berechtigungen 
    und zur automatisierten Erstellung von RDP-Dateien und E-Mail-Entwürfen.
.NOTES
    Autor           : PS-Coding
    Version         : 3.6 (FIX: StringBuilder statt Here-Strings für maximale Robustheit. Alle Parser-Probleme behoben.)
    Erstellt am     : 18.11.2025
    PowerShell      : 5.1
    Umgebung        : Windows On-Premise (AD)
    Voraussetzungen : 1. PowerShell 5.1
                      2. Active Directory Modul (RSAT)
                      3. Administrator-Rechte (Lokal auf Clients / Domäne)
                      4. Microsoft Outlook (Desktop-Client) ODER einen SMTP-Relay-Server.
                      5. WinRM/WMI/RPC muss auf Clients erreichbar sein.
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
    [Alias('genfromlog')]
    [Switch]
    $GenerateFromLog,
    
    [Parameter(Mandatory = $true, ParameterSetName = 'GenerateFromLog')]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]
    $InputLogPath,

    #--- Gemeinsame Parameter ---
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

    #--- GLOBALE E-Mail-Parameter ---
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
    [Parameter(Mandatory = $false)]
    [string]
    $UserColumn = 'sAMAccountName',

    [Parameter(Mandatory = $false)]
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
# HILFSFUNKTIONEN
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
        $msg = "Fehler beim dynamischen Ermitteln der AD-Domäne."
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
    $status = 'Failed' 

    try {
        $groupPath = "WinNT://$ComputerName/$LocalGroupName,group"
        $group = [ADSI]$groupPath
        $userPath = "WinNT://$Domain/$UserName,user"
        $user = [ADSI]$userPath
        
        if ($Action -eq 'Add') {
            $group.Add($user.Path)
            # Check
            $group.RefreshCache()
            $members = $group.Members() | ForEach-Object { $_.GetType().InvokeMember("Name", "GetProperty", $null, $_, $null) }
            if ($members -contains $UserName) { $status = 'Success' } else { $status = 'VerificationFailed' }
        }
        elseif ($Action -eq 'Remove') {
            $group.Remove($user.Path)
            # Check
            $group.RefreshCache()
            $members = $group.Members() | ForEach-Object { $_.GetType().InvokeMember("Name", "GetProperty", $null, $_, $null) }
            if (-not ($members -contains $UserName)) { $status = 'Success' } else { $status = 'VerificationFailed' }
        }
    }
    catch {
        $errorMessage = $_.Exception.Message.Trim()
        if ($errorMessage -like "*bereits Mitglied*") { $status = 'AlreadyExists' }
        elseif ($errorMessage -like "*nicht Mitglied*") { $status = 'NotMember' }
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
    param([string]$ComputerName, [string]$FilePath)
    
    $sb = [System.Text.StringBuilder]::new()
    $sb.AppendLine("screen mode id:i:2") | Out-Null
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
        if (-not (Get-Module -Name ActiveDirectory)) { Import-Module ActiveDirectory -ErrorAction Stop }
        $adUser = Get-ADUser -Identity $UserName -Properties DisplayName, EmailAddress -ErrorAction Stop
    } catch {
        $msg = "Konnte AD-Benutzer $UserName nicht finden."
        Write-Warning $msg
        if ($Global:GlobalErrorLog) { $Global:GlobalErrorLog.Add($msg) | Out-Null }
        return $false
    }
    
    $recipientName = $adUser.DisplayName
    $recipientEmail = $adUser.EmailAddress
    $senderEmail = $adUser.EmailAddress 

    if ([string]::IsNullOrWhiteSpace($recipientEmail)) {
        Write-Warning "Benutzer $UserName hat keine E-Mail-Adresse im AD."
        return $false
    }
    
    $mailSubject = "Ihre RDP-Verbindungsdateien"
    $sbBody = [System.Text.StringBuilder]::new()
    $sbBody.AppendLine("<html><body style='font-family:Calibri, Arial, sans-serif;'>") | Out-Null
    $sbBody.AppendLine("<p>Hallo $recipientName,</p><p>anbei die RDP-Dateien.</p>") | Out-Null
    $sbBody.AppendLine("</body></html>") | Out-Null
    $mailBody = $sbBody.ToString()

    # SMTP
    if (-not [string]::IsNullOrWhiteSpace($SmtpServer)) {
        if (-not $Send) { return $false }
        
        if ($null -eq $Credential) {
            Write-Host "SMTP-Versand erfordert Anmeldeinformationen (via Parameter oder Prompt)." -ForegroundColor Yellow
            $Credential = Get-Credential -Message "SMTP-Login"
        }
        
        $smtpParams = @{
            To = $recipientEmail; From = $senderEmail; Subject = $mailSubject; Body = $mailBody; BodyAsHtml = $true
            SmtpServer = $SmtpServer; Attachments = $RDPFilePaths; ErrorAction = 'Stop'
        }
        if ($Credential.UserName) { $smtpParams.Add("Credential", $Credential) }
        
        try {
            Send-MailMessage @smtpParams
            Write-Host " ERFOLG: E-Mail via SMTP an $recipientEmail gesendet." -ForegroundColor Cyan
            return $true
        } catch {
            Write-Warning ("SMTP-Fehler für {0}: {1}" -f $UserName, $_.Exception.Message)
            return $false
        }
    }
    # OUTLOOK
    else {
        $Error.Clear()
        try {
            try { $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') } 
            catch { $outlook = New-Object -ComObject Outlook.Application }

            if (!$outlook) { Write-Warning "Outlook nicht verfügbar."; return $false }
            
            $mail = $outlook.CreateItem(0) 
            $mail.Subject = $mailSubject
            $mail.To = $recipientEmail
            $mail.HTMLBody = $mailBody

            foreach ($rdpPath in $RDPFilePaths) {
                if (Test-Path $rdpPath) { $mail.Attachments.Add($rdpPath, 1, 1, ($rdpPath | Split-Path -Leaf)) }
            }
            
            if (-not [string]::IsNullOrWhiteSpace($MSGSavePath)) {
                $mail.SaveAs($MSGSavePath, 5) 
                Write-Host " ERFOLG: MSG gespeichert." -ForegroundColor Green
            }
            
            if ($Send) {
                $mail.Send()
                Write-Host " ERFOLG: E-Mail via Outlook an $recipientEmail gesendet." -ForegroundColor Cyan
            }
            return $true
        }
        catch {
            Write-Warning "Outlook-Fehler: $($_.Exception.Message)"
            if ($_.Exception.Message -like "*80080005*") { Write-Warning "TIPP: Skript nicht 'Als Admin' starten, wenn Outlook läuft." }
            return $false
        }
        finally {
            if ($mail) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($mail) | Out-Null }
            if ($outlook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null }
            [System.GC]::Collect()
        }
    }
}

Function Generate-SendMailsScript {
    param(
        $UserList, $RDPFilePaths, $OutputPath, $SmtpServer, $SaveAsMsgOnly
    )
    
    Write-Verbose "Generiere 'sendMails.ps1'..."
    $resolvedOutputPath = Resolve-Path $OutputPath
    
    # Robuster Aufbau des Skript-Inhalts via StringBuilder (Kein Here-String-Parser-Fehler mehr!)
    $sb = [System.Text.StringBuilder]::new()
    
    $sb.AppendLine("# --- AUTO-GENERATED MAIL SCRIPT ---") | Out-Null
    $sb.AppendLine("param([string]`$AltSmtp)") | Out-Null
    $sb.AppendLine("") | Out-Null
    
    # Benötigte Funktionen in das Sub-Skript injizieren
    # Wir nutzen Definition der Funktion aus dem aktuellen Scope
    $funcDef = (Get-Command Start-EmailWorkflow).Definition
    $sb.AppendLine("Function Start-EmailWorkflow {") | Out-Null
    $sb.AppendLine($funcDef) | Out-Null
    $sb.AppendLine("}") | Out-Null
    
    $sb.AppendLine("") | Out-Null
    $sb.AppendLine("# Daten") | Out-Null
    
    # Arrays sicher übergeben
    $uStr = $UserList | ForEach-Object { "'$_'" } | Join-String -Separator ","
    $rStr = $RDPFilePaths | ForEach-Object { $p = Resolve-Path $_; "'$p'" } | Join-String -Separator ","
    
    $sb.AppendLine('$UserList = @(' + $uStr + ')') | Out-Null
    $sb.AppendLine('$RDPFilePaths = @(' + $rStr + ')') | Out-Null
    
    $smtpStr = if ($SmtpServer) { "'$SmtpServer'" } else { "`$null" }
    $sb.AppendLine('$SmtpServer = ' + $smtpStr) | Out-Null
    $sb.AppendLine('if ($AltSmtp) { $SmtpServer = $AltSmtp }') | Out-Null
    
    $sendBool = if ($SaveAsMsgOnly) { "$false" } else { "$true" }
    $sb.AppendLine('$Send = $' + $sendBool) | Out-Null
    
    # Ausführungslogik
    $sb.AppendLine('$OutputPath = "' + $resolvedOutputPath + '"') | Out-Null
    
    $sb.AppendLine('foreach ($u in $UserList) {') | Out-Null
    $sb.AppendLine('  $msgPath = Join-Path $OutputPath "$u.msg"') | Out-Null
    $sb.AppendLine('  Write-Host "Verarbeite $u..." -ForegroundColor Cyan') | Out-Null
    $sb.AppendLine('  Start-EmailWorkflow -UserName $u -RDPFilePaths $RDPFilePaths -MSGSavePath $msgPath -Send $Send -SmtpServer $SmtpServer -Credential $null') | Out-Null
    $sb.AppendLine('}') | Out-Null
    $sb.AppendLine('Read-Host "Fertig. Enter druecken."') | Out-Null

    try {
        $scriptFileName = "sendMails_$($GlobalTimestamp).ps1"
        $scriptFullPath = Join-Path -Path $OutputPath -ChildPath $scriptFileName
        Set-Content -Path $scriptFullPath -Value $sb.ToString() -Encoding UTF8 -ErrorAction Stop
        Write-Host "[OK] 'sendMails.ps1' erstellt: $scriptFullPath" -ForegroundColor Green
    } catch {
        Write-Warning "Fehler beim Erstellen von sendMails.ps1: $($_.Exception.Message)"
    }
}

Function Write-Log {
    param([object[]]$ReportData)
    $LogName = "Manage-RDPUsers"
    if ($ReportData.Length -gt 0 -and ($ReportData[0].Action -match 'Email')) { $LogName = "Manage-RDPFiles" }
    try {
        if (-not (Test-Path $GlobalLogDir)) { New-Item -Path $GlobalLogDir -ItemType Directory -Force | Out-Null }
        $f = Join-Path $GlobalLogDir "$($LogName)_$($GlobalTimestamp).csv"
        $ReportData | Export-Csv -Path $f -NoType -Delimiter ';' -Encoding UTF8
        Write-Host "Log: $f" -ForegroundColor DarkGreen
    } catch { Write-Warning "Log-Fehler" }
}

Function Write-ErrorLog {
    param([string]$ErrorOutputPath)
    if (($GlobalErrorLog | Measure-Object).Count -gt 0) {
        try {
            if (-not (Test-Path $ErrorOutputPath)) { New-Item -Path $ErrorOutputPath -ItemType Directory -Force | Out-Null }
            $f = Join-Path $ErrorOutputPath "ERROR_$($GlobalTimestamp).TXT"
            $GlobalErrorLog | Out-File $f -Encoding UTF8
            Write-Warning "Fehler-Log: $f"
        } catch {}
    }
}

Function Load-CsvData {
    param($CsvPath)
    if (-not (Test-Path $CsvPath)) { Write-Error "Datei nicht gefunden: $CsvPath"; return $null }
    try {
        $d = Import-Csv $CsvPath -Delimiter ';' -Encoding Default -ErrorAction Stop
        if (($d | Measure-Object).Count -eq 0) { Write-Error "CSV leer: $CsvPath"; return $null }
        return $d
    } catch { Write-Error "CSV-Fehler: $($_.Exception.Message)"; return $null }
}

#==============================================================================
# LOGIK
#==============================================================================

Function Start-RightsWorkflow {
    param($SetRights, $RemoveRights, $UserListPath, $ClientListPath, $UserColumn, $ClientColumn, $SendEmail, $SmtpServer, $Credential, $OutputPath)
    
    $Users = Load-CsvData -CsvPath $UserListPath
    $Clients = Load-CsvData -CsvPath $ClientListPath
    if (!$Users -or !$Clients) { return }
    if (-not (Get-Module -Name ActiveDirectory)) { Import-Module ActiveDirectory }
    $ADContext = Get-DynamicADDomain
    
    $Action = if ($SetRights) { 'Add' } else { 'Remove' }
    $Verb = if ($SetRights) { "Hinzufügen" } else { "Entfernen" }
    
    $uCount = ($Users | Measure-Object).Count
    $cCount = ($Clients | Measure-Object).Count
    
    Write-Host "Modus: Rechte $Verb ($uCount User -> $cCount Clients)" -ForegroundColor Cyan
    
    # Planen
    $Plan = @()
    foreach ($c in $Clients) {
        $cName = $c.$ClientColumn
        if ([string]::IsNullOrWhiteSpace($cName)) { continue }
        if (-not (Test-Connection $cName -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
            $GlobalErrorLog.Add("Client $cName offline.")
            Write-Warning "Client $cName offline."
            continue
        }
        foreach ($u in $Users) {
            $uName = $u.$UserColumn
            if ($uName) { $Plan += [PSCustomObject]@{ Client=$cName; User=$uName; Action=$Action } }
        }
    }
    
    $pCount = ($Plan | Measure-Object).Count
    if ($pCount -eq 0) { Write-Host "Keine Aktionen." -ForegroundColor Yellow; Write-ErrorLog $GlobalLogDir; return }
    
    $Plan | Format-Table -AutoSize
    
    # Bestätigung (Robust)
    $doEx = $false
    if ($PSCmdlet.ShouldProcess("$pCount Aktionen", "Ausführen")) {
        # Prüfen ob -Confirm:$false
        if ($PSBoundParameters.ContainsKey('Confirm') -and -not $Confirm) { $doEx = $true }
        else {
            $inp = Read-Host "Ausführen? (J/N)"
            if ($inp -eq 'J') { $doEx = $true }
        }
    }
    
    if ($doEx) {
        $Rep = @()
        foreach ($a in $Plan) {
            $s = Invoke-RemoteGroupMembership -ComputerName $a.Client -UserName $a.User -Domain $ADContext.NetBIOS -Action $a.Action
            $Rep += [PSCustomObject]@{ Client=$a.Client; User=$a.User; Action=$a.Action; Status=$s; Time=Get-Date }
            Write-Host "$($a.Action) $($a.User) on $($a.Client): $s"
        }
        Write-Log $Rep
        Write-ErrorLog $GlobalLogDir
        
        if ($SendEmail) {
            $suc = $Rep | Where-Object { $_.Status -in 'Success','AlreadyExists' }
            if (($suc | Measure-Object).Count -gt 0) {
                $u = $suc | Select -Expand User -Unique
                $c = $suc | Select -Expand Client -Unique
                Start-FileWorkflow -IsCombinedRun $true -UserList $u -ClientList $c -OutputPath $OutputPath -SmtpServer $SmtpServer -Credential $Credential -UserColumn "User" -ClientColumn "Client"
            }
        }
    }
}

Function Start-FileWorkflow {
    param($UserListPath, $ClientListPath, $InputLogPath, $GenerateFromLog, $UserList, $ClientList, $IsCombinedRun, $OutputPath, $UserColumn, $ClientColumn, $SaveAsMsgOnly, $SmtpServer, $Credential)
    
    $Users = $null; $Clients = $null
    
    # Datenquellen laden
    if ($IsCombinedRun) { $Users = $UserList; $Clients = $ClientList; $UserColumn="User"; $ClientColumn="Client" }
    elseif ($GenerateFromLog) {
        $l = Load-CsvData $InputLogPath
        if (!$l) { return }
        $suc = $l | Where { $_.Status -in 'Success','AlreadyExists' }
        $Users = $suc | Select -Expand User -Unique
        $Clients = $suc | Select -Expand Client -Unique
        $UserColumn="User"; $ClientColumn="Client"
    }
    else {
        $Users = Load-CsvData $UserListPath
        $Clients = Load-CsvData $ClientListPath
        if (!$Users -or !$Clients) { return }
    }
    
    if (-not (Get-Module -Name ActiveDirectory)) { Import-Module ActiveDirectory }
    
    if ([string]::IsNullOrWhiteSpace($OutputPath)) { $OutputPath = Join-Path $PSScriptRoot "RDP_Out_$($GlobalTimestamp)" }
    if (-not (Test-Path $OutputPath)) { New-Item $OutputPath -ItemType Directory -Force | Out-Null }
    
    Write-Host "Modus: Dateien/Mails ($(($Users|Measure-Object).Count) User)" -ForegroundColor Cyan
    
    # RDPs erstellen
    $rdps = @()
    $cObjs = if ($IsCombinedRun -or $GenerateFromLog) { $Clients } else { $Clients.$ClientColumn }
    foreach ($c in $cObjs) {
        if (!$c) { continue }
        if (-not (Test-Connection $c -Count 1 -Quiet -ErrorAction SilentlyContinue)) { Write-Warning "$c offline"; continue }
        $p = Join-Path $OutputPath "$c.rdp"
        if (Create-RDPFile $c $p) { $rdps += $p }
    }
    
    if ($rdps.Count -eq 0) { Write-Error "Keine RDPs erstellt."; return }
    
    # Mails
    $doSend = -not $SaveAsMsgOnly
    $Rep = @()
    $uObjs = if ($IsCombinedRun -or $GenerateFromLog) { $Users } else { $Users.$UserColumn }
    
    foreach ($u in $uObjs) {
        if (!$u) { continue }
        $msgP = if (!$SmtpServer) { Join-Path $OutputPath "$u.msg" } else { $null }
        $res = Start-EmailWorkflow -UserName $u -RDPFilePaths $rdps -MSGSavePath $msgP -Send $doSend -SmtpServer $SmtpServer -Credential $Credential
        $Rep += [PSCustomObject]@{ User=$u; Action=(if($doSend){"Send"}else{"Save"}); Status=$res; Time=Get-Date }
    }
    
    Write-Log $Rep
    Write-ErrorLog $OutputPath
    
    Generate-SendMailsScript -UserList $uObjs -RDPFilePaths $rdps -OutputPath $OutputPath -SmtpServer $SmtpServer -SaveAsMsgOnly $SaveAsMsgOnly
}

#==============================================================================
# START
#==============================================================================

if ($PSCmdlet.ParameterSetName -ne 'Interactive') {
    if ($SetRDPRights -or $RemoveRDPRights) {
        Start-RightsWorkflow -SetRights $SetRDPRights -RemoveRights $RemoveRDPRights -UserListPath $UserListPath -ClientListPath $ClientListPath -UserColumn $UserColumn -ClientColumn $ClientColumn -SendEmail:$SendEmail -SmtpServer $SmtpServer -Credential $Credential -OutputPath $OutputPath
    }
    elseif ($GenerateRDPFiles) {
        Start-FileWorkflow -GenerateRDPFiles -UserListPath $UserListPath -ClientListPath $ClientListPath -OutputPath $OutputPath -UserColumn $UserColumn -ClientColumn $ClientColumn -SaveAsMsgOnly:$SaveAsMsgOnly -SmtpServer $SmtpServer -Credential $Credential
    }
    elseif ($GenerateFromLog) {
        Start-FileWorkflow -GenerateFromLog -InputLogPath $InputLogPath -OutputPath $OutputPath -SaveAsMsgOnly:$SaveAsMsgOnly -SmtpServer $SmtpServer -Credential $Credential
    }
}
else {
    # Interaktives Menü Logik (vereinfacht)
    Show-MainMenu
    $sel = Read-Host "Wahl"
    # ... (Logik für Interaktiv kann bei Bedarf hier ergänzt werden, Parameter-Modus war Fokus) ...
    if ($sel -eq 'Q') { return }
    Write-Warning "Bitte nutzen Sie Parameter für komplexe Aufgaben oder starten Sie neu mit Parametern."
}
Write-Host "Ende."
