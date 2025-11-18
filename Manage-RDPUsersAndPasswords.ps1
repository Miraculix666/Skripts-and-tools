<#
.FILENAME Manage-RDPUsers.ps1
.DESCRIPTION
    Ein PowerShell-Skript (v5.1) zur Verwaltung von Remotedesktop-Berechtigungen 
    und zur automatisierten Erstellung von RDP-Dateien und E-Mail-Entwürfen.
.NOTES
    Autor           : PS-Coding
    Version         : 3.7 (FIX: Here-Strings in Generate-SendMailsScript durch StringBuilder ersetzt. Parser-Fehler endgültig behoben.)
    Erstellt am     : 18.11.2025
    PowerShell      : 5.1
    Umgebung        : Windows On-Premise (AD)
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
        $msg = "Benutzer $UserName hat keine E-Mail-Adresse im AD."
        Write-Warning $msg
        if ($Global:GlobalErrorLog) { $Global:GlobalErrorLog.Add($msg) | Out-Null }
        return $false
    }
    
    $mailSubject = "Ihre RDP-Verbindungsdateien"
    
    # String Builder für Mail Body (Robust)
    $sbBody = [System.Text.StringBuilder]::new()
    $sbBody.AppendLine("<html><body style='font-family:Calibri, Arial, sans-serif;'>") | Out-Null
    $sbBody.AppendLine("<p>Hallo $recipientName,</p>") | Out-Null
    $sbBody.AppendLine("<p>anbei erhalten Sie Ihre RDP-Zugriffsdateien.</p>") | Out-Null
    $sbBody.AppendLine("<p><i>Automatisch generiert</i></p>") | Out-Null
    $sbBody.AppendLine("</body></html>") | Out-Null
    $mailBody = $sbBody.ToString()

    # SMTP
    if (-not [string]::IsNullOrWhiteSpace($SmtpServer)) {
        if (-not $Send) { return $false }
        
        if ($null -eq $Credential) {
            Write-Host "SMTP-Versand erfordert Anmeldeinformationen." -ForegroundColor Yellow
            # Hier kein Read-Host, Credentials müssen übergeben worden sein
            return $false
        }
        
        $smtpParams = @{
            To = $recipientEmail; From = $senderEmail; Subject = $mailSubject; Body = $mailBody; BodyAsHtml = $true
            SmtpServer = $SmtpServer; Attachments = $RDPFilePaths; ErrorAction = 'Stop'
        }
        if ($Credential.UserName) { $smtpParams.Add("Credential", $Credential) }
        
        try {
            Send-MailMessage @smtpParams
            Write-Host " [SMTP] Gesendet an $recipientEmail" -ForegroundColor Cyan
            return $true
        } catch {
            $errMsg = $_.Exception.Message
            Write-Warning ("SMTP-Fehler für {0}: {1}" -f $UserName, $errMsg)
            if ($Global:GlobalErrorLog) { $Global:GlobalErrorLog.Add("SMTP Fehler $UserName: $errMsg") | Out-Null }
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
                Write-Host " [MSG] Gespeichert: $MSGSavePath" -ForegroundColor Green
            }
            
            if ($Send) {
                $mail.Send()
                Write-Host " [OUTLOOK] Gesendet an $recipientEmail" -ForegroundColor Cyan
            }
            return $true
        }
        catch {
            $errMsg = $_.Exception.Message
            Write-Warning "Outlook-Fehler: $errMsg"
            if ($errMsg -like "*80080005*") { 
                Write-Warning "TIPP: Skript nicht 'Als Admin' starten, wenn Outlook als User läuft. Nutzen Sie 'sendMails.ps1'." 
            }
            if ($Global:GlobalErrorLog) { $Global:GlobalErrorLog.Add("Outlook Fehler $UserName: $errMsg") | Out-Null }
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
    param($UserList, $RDPFilePaths, $OutputPath, $SmtpServer, $SaveAsMsgOnly)
    
    Write-Verbose "Generiere 'sendMails.ps1'..."
    $resOut = Resolve-Path $OutputPath
    
    # --- FIX V3.7: StringBuilder statt Here-Strings für das gesamte Skript ---
    $sb = [System.Text.StringBuilder]::new()
    
    $sb.AppendLine("# --- AUTO-GENERATED MAIL SCRIPT ---") | Out-Null
    $sb.AppendLine("param([string]`$AltSmtp)") | Out-Null
    $sb.AppendLine("") | Out-Null
    
    # Funktion injizieren
    $funcDef = (Get-Command Start-EmailWorkflow).Definition
    $sb.AppendLine("Function Start-EmailWorkflow {") | Out-Null
    $sb.AppendLine($funcDef) | Out-Null
    $sb.AppendLine("}") | Out-Null
    $sb.AppendLine("") | Out-Null
    
    # Daten injizieren
    $uStr = $UserList | ForEach-Object { "'$_'" } | Join-String -Separator ","
    $rStr = $RDPFilePaths | ForEach-Object { $p = Resolve-Path $_; "'$p'" } | Join-String -Separator ","
    $sb.AppendLine('$UserList = @(' + $uStr + ')') | Out-Null
    $sb.AppendLine('$RDPFilePaths = @(' + $rStr + ')') | Out-Null
    
    $smtpStr = if ($SmtpServer) { "'$SmtpServer'" } else { "`$null" }
    $sb.AppendLine('$SmtpServer = ' + $smtpStr) | Out-Null
    $sb.AppendLine('if ($AltSmtp) { $SmtpServer = $AltSmtp }') | Out-Null
    
    $sendBool = if ($SaveAsMsgOnly) { "$false" } else { "$true" }
    $sb.AppendLine('$Send = $' + $sendBool) | Out-Null
    $sb.AppendLine('$OutputPath = "' + $resOut + '"') | Out-Null
    
    # Ausführungsschleife injizieren
    $sb.AppendLine("") | Out-Null
    $sb.AppendLine('foreach ($u in $UserList) {') | Out-Null
    $sb.AppendLine('  $msgPath = Join-Path $OutputPath "$u.msg"') | Out-Null
    $sb.AppendLine('  if ($SmtpServer) { $msgPath = $null }') | Out-Null
    $sb.AppendLine('  Write-Host "Verarbeite $u..." -ForegroundColor Cyan') | Out-Null
    $sb.AppendLine('  Start-EmailWorkflow -UserName $u -RDPFilePaths $RDPFilePaths -MSGSavePath $msgPath -Send $Send -SmtpServer $SmtpServer -Credential $null') | Out-Null
    $sb.AppendLine('}') | Out-Null
    $sb.AppendLine('Read-Host "Fertig. Enter druecken."') | Out-Null

    try {
        $scriptFileName = "sendMails_$($GlobalTimestamp).ps1"
        $scriptFullPath = Join-Path -Path $OutputPath -ChildPath $scriptFileName
        Set-Content -Path $scriptFullPath -Value $sb.ToString() -Encoding UTF8 -ErrorAction Stop
        
        Write-Host "-------------------------------------------------------------" -ForegroundColor Cyan
        Write-Host "[OK] 'sendMails.ps1' erstellt: $scriptFullPath" -ForegroundColor Green
        Write-Host "     (Nutzen Sie dies, falls der direkte Outlook-Versand fehlschlaegt)"
        Write-Host "-------------------------------------------------------------" -ForegroundColor Cyan
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
        Write-Host " [LOG] $f" -ForegroundColor DarkGreen
    } catch { Write-Warning "Log-Fehler" }
}

Function Write-ErrorLog {
    param([string]$ErrorOutputPath)
    if (($GlobalErrorLog | Measure-Object).Count -gt 0) {
        try {
            if (-not (Test-Path $ErrorOutputPath)) { New-Item -Path $ErrorOutputPath -ItemType Directory -Force | Out-Null }
            $f = Join-Path $ErrorOutputPath "ERROR_$($GlobalTimestamp).TXT"
            $GlobalErrorLog | Out-File $f -Encoding UTF8
            Write-Warning "Fehler aufgetreten. Details: $f"
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
# WORKFLOWS
#==============================================================================

Function Start-RightsWorkflow {
    param($Set, $UserListPath, $ClientListPath, $UserCol, $ClientCol, $DoMail, $Smtp, $Cred, $OutPath)

    $Users = Load-CsvData $UserListPath
    $Clients = Load-CsvData $ClientListPath
    if (!$Users -or !$Clients) { return }

    if (-not (Get-Module -Name ActiveDirectory)) { try{Import-Module ActiveDirectory -ErrorAction Stop}catch{Write-Error "AD Modul fehlt";return} }
    $AD = Get-DynamicADDomain
    
    $Act = if ($Set) { 'Add' } else { 'Remove' }
    $Verb = if ($Set) { "Hinzufügen" } else { "Entfernen" }

    Write-Host "--- MODUS: Rechte $Verb ---" -ForegroundColor Cyan
    Write-Host " Benutzer: $(($Users|Measure-Object).Count)"
    Write-Host " Clients:  $(($Clients|Measure-Object).Count)"
    
    # Planung
    Write-Host " [1/3] Prüfe Erreichbarkeit und plane Aktionen..." -ForegroundColor Yellow
    $Plan = @()
    
    foreach ($c in $Clients) {
        $cN = $c.$ClientCol
        if ([string]::IsNullOrWhiteSpace($cN)) { continue }
        if (Test-Connection $cN -Count 1 -Quiet) {
            foreach ($u in $Users) {
                if ($u.$UserCol) { $Plan += [PSCustomObject]@{ Client=$cN; User=$u.$UserCol; Action=$Act } }
            }
        } else {
            Write-Warning "Client $cN offline."
            $GlobalErrorLog.Add("Client $cN offline") | Out-Null
        }
    }
    
    $count = ($Plan | Measure-Object).Count
    if ($count -eq 0) { Write-Warning "Keine Aktionen möglich."; Write-ErrorLog $GlobalLogDir; return }
    
    # Bestätigung
    Write-Host " [2/3] $count Aktionen geplant." -ForegroundColor Yellow
    $Plan | Format-Table -AutoSize
    
    if ($PSCmdlet.ShouldProcess("$count Aktionen", "Ausführen")) {
        if ($PSBoundParameters.ContainsKey('Confirm') -and -not $Confirm) { 
            # Force exec
        } else {
            $in = Read-Host " Starten? (J/N)"
            if ($in -ne 'J') { Write-Warning "Abbruch."; return }
        }
    } else { return } # WhatIf

    # Ausführung
    Write-Host " [3/3] Ausführung..." -ForegroundColor Cyan
    $Rep = @()
    foreach ($p in $Plan) {
        $res = Invoke-RemoteGroupMembership -ComputerName $p.Client -UserName $p.User -Domain $AD.NetBIOS -Action $p.Action
        
        $col = if ($res -eq 'Success') { "Green" } elseif ($res -match 'Already|NotMember') { "Gray" } else { "Red" }
        Write-Host " $res : $($p.User) -> $($p.Client)" -ForegroundColor $col
        
        $Rep += [PSCustomObject]@{ Client=$p.Client; User=$p.User; Action=$p.Action; Status=$res; Time=Get-Date }
    }
    
    Write-Log -Data $Rep -Type "Rights"
    Write-ErrorLog $GlobalLogDir
    
    # Mail-Chaining
    if ($DoMail) {
        $suc = $Rep | Where { $_.Status -in 'Success','AlreadyExists' }
        if (($suc | Measure-Object).Count -gt 0) {
            $uList = $suc | Select -Expand User -Unique
            $cList = $suc | Select -Expand Client -Unique
            Start-FileWorkflow -UserList $uList -ClientList $cList -IsCombined $true -OutputPath $OutPath -Smtp $Smtp -Cred $Cred -UserCol "User" -ClientCol "Client"
        } else {
            Write-Warning "Keine erfolgreichen Rechtezuweisungen für Mail-Versand."
        }
    }
}

Function Start-FileWorkflow {
    param($UserListPath, $ClientListPath, $LogPath, $FromLog, $UserList, $ClientList, $IsCombined, $OutputPath, $UserCol, $ClientCol, $SaveOnly, $Smtp, $Cred)

    $Users = $null; $Clients = $null
    
    if ($IsCombined) { $Users = $UserList; $Clients = $ClientList; $UserCol="User"; $ClientCol="Client" }
    elseif ($FromLog) {
        $l = Load-CsvData $LogPath
        if (!$l) { return }
        $suc = $l | Where { $_.Status -in 'Success','AlreadyExists' }
        $Users = $suc | Select -Expand User -Unique
        $Clients = $suc | Select -Expand Client -Unique
        $UserCol="User"; $ClientCol="Client"
    }
    else {
        $Users = Load-CsvData $UserListPath
        $Clients = Load-CsvData $ClientListPath
        if (!$Users -or !$Clients) { return }
    }
    
    if (-not (Get-Module -Name ActiveDirectory)) { Import-Module ActiveDirectory }
    
    if ([string]::IsNullOrWhiteSpace($OutputPath)) { $OutputPath = Join-Path $PSScriptRoot "RDP_Out_$($GlobalTimestamp)" }
    if (-not (Test-Path $OutputPath)) { New-Item $OutputPath -ItemType Directory -Force | Out-Null }
    
    Write-Host "--- MODUS: Dateien erstellen ---" -ForegroundColor Cyan
    
    # RDPs
    $rdps = @()
    $cObjs = if ($IsCombined -or $FromLog) { $Clients } else { $Clients.$ClientCol }
    
    foreach ($cn in $cObjs) {
        if ([string]::IsNullOrWhiteSpace($cn)) { continue }
        if (-not (Test-Connection $cn -Count 1 -Quiet -ErrorAction SilentlyContinue)) { Write-Warning "$cn offline"; continue }
        $p = Join-Path $OutputPath "$cn.rdp"
        if (Create-RDPFile $cn $p) { $rdps += $p }
    }
    
    if ($rdps.Count -eq 0) { Write-Error "Keine RDPs erstellt."; return }
    Write-Host " $(($rdps|Measure-Object).Count) RDP-Dateien erstellt." -ForegroundColor Green
    
    # Mails
    $doSend = -not $SaveOnly
    $mRep = @()
    $uObjs = if ($IsCombined -or $FromLog) { $Users } else { $Users.$UserCol }
    
    Write-Host " Verarbeite E-Mails..."
    foreach ($un in $uObjs) {
        if ([string]::IsNullOrWhiteSpace($un)) { continue }
        $mp = if (!$Smtp) { Join-Path $OutputPath "$un.msg" } else { $null }
        $ok = Start-EmailWorkflow -UserName $un -RDPFilePaths $rdps -MSGSavePath $mp -Send $doSend -SmtpServer $Smtp -Credential $Cred
        $mRep += [PSCustomObject]@{ User=$un; Action=(if($doSend){"Send"}else{"Save"}); Status=$ok; Time=Get-Date }
    }
    
    Write-Log -Data $mRep -Type "Mails"
    Write-ErrorLog $OutputPath
    
    Generate-SendMailsScript -UserList $uObjs -RDPFilePaths $rdps -OutputPath $OutputPath -SmtpServer $Smtp -SaveAsMsgOnly $SaveOnly
}

#==============================================================================
# MAIN
#==============================================================================

if ($PSCmdlet.ParameterSetName -ne 'Interactive') {
    if ($SetRDPRights -or $RemoveRDPRights) {
        Start-RightsWorkflow -Set $SetRDPRights -UserListPath $UserListPath -ClientListPath $ClientListPath -UserCol $UserColumn -ClientCol $ClientColumn -DoMail $SendEmail -Smtp $SmtpServer -Cred $Credential -OutPath $OutputPath
    }
    elseif ($GenerateRDPFiles) {
        Start-FileWorkflow -UserListPath $UserListPath -ClientListPath $ClientListPath -OutputPath $OutputPath -UserCol $UserColumn -ClientCol $ClientColumn -SaveOnly $SaveAsMsgOnly -Smtp $SmtpServer -Cred $Credential
    }
    elseif ($GenerateFromLog) {
        Start-FileWorkflow -FromLog $true -LogPath $InputLogPath -OutputPath $OutputPath -SaveOnly $SaveAsMsgOnly -Smtp $SmtpServer -Cred $Credential
    }
}
else {
    Show-MainMenu
    $sel = Read-Host "Wahl"
    if ($sel -eq '1') {
        $u = Read-Host "User-CSV Pfad"; $c = Read-Host "Client-CSV Pfad"
        Start-RightsWorkflow -Set $true -UserListPath $u -ClientListPath $c
    }
    if ($sel -eq '3') {
        $u = Read-Host "User-CSV Pfad"; $c = Read-Host "Client-CSV Pfad"
        Start-FileWorkflow -UserListPath $u -ClientListPath $c
    }
    # (Vereinfachtes Menü für Stabilität)
    if ($sel -eq 'Q') { return }
}
Write-Host "Ende."
