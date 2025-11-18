<#
.SYNOPSIS
    RDP User Manager & Mail Generator (v4.0)

.DESCRIPTION
    Dieses Skript verwaltet die lokale Gruppe "Remotedesktopbenutzer" auf entfernten Windows-Clients
    (Hinzufügen/Entfernen von Domänen-Benutzern) und generiert automatisiert RDP-Dateien sowie
    E-Mail-Benachrichtigungen (Outlook oder SMTP).

    Features:
    - Massenverwaltung von Rechten (CSV-basiert)
    - Prüfung der Erreichbarkeit (Ping) vor der Ausführung
    - Sicherheitsabfrage (J/N) mit Vorschau
    - Automatische AD-Abfrage für E-Mail-Adressen und Namen
    - Generierung von 'sendMails.ps1' zur Umgehung von Admin/User-Kontext-Problemen bei Outlook
    - Detailliertes Logging (CSV) und Fehlerprotokollierung (ERROR.TXT)

.PARAMETER SetRDPRights
    Modus 1: Fügt Benutzer zur Gruppe hinzu.

.PARAMETER RemoveRDPRights
    Modus 2: Entfernt Benutzer aus der Gruppe.

.PARAMETER GenerateRDPFiles
    Modus 3: Erstellt RDP-Dateien und E-Mails (basierend auf CSV).
    Alias: generatemail

.PARAMETER GenerateFromLog
    Modus 4: Wiederholt den E-Mail-Workflow basierend auf einem erfolgreichen Log-File.
    Alias: genfromlog

.PARAMETER UserListPath
    Pfad zur CSV-Datei mit Benutzernamen (Header: sAMAccountName).

.PARAMETER ClientListPath
    Pfad zur CSV-Datei mit Computernamen (Header: ComputerName).

.PARAMETER InputLogPath
    Pfad zur Log-Datei für Modus 4.

.PARAMETER SendEmail
    Optional für Modus 1/2: Startet nach der Rechtevergabe sofort den E-Mail-Workflow.
    Standard für Modus 3/4: Aktiviert.

.PARAMETER SaveAsMsgOnly
    Verhindert den direkten Versand und speichert nur .msg-Entwürfe.

.PARAMETER SmtpServer
    Verwendet SMTP statt Outlook.

.NOTES
    Version:    4.0 (Final Polish)
    Autor:      PS-Coding
    Datum:      18.11.2025
#>

[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'Interactive')]
param(
    #--- MODUS-SCHALTER ---
    [Parameter(Mandatory = $false, ParameterSetName = 'Set-RDPRights')]
    [Switch]$SetRDPRights,

    [Parameter(Mandatory = $false, ParameterSetName = 'Remove-RDPRights')]
    [Switch]$RemoveRDPRights,

    [Parameter(Mandatory = $false, ParameterSetName = 'Generate-RDPFiles')]
    [Alias('generatemail')]
    [Switch]$GenerateRDPFiles,

    [Parameter(Mandatory = $false, ParameterSetName = 'GenerateFromLog')]
    [Alias('genfromlog')]
    [Switch]$GenerateFromLog,
    
    #--- EINGABE-DATEIEN ---
    [Parameter(Mandatory = $true, ParameterSetName = 'Set-RDPRights')]
    [Parameter(Mandatory = $true, ParameterSetName = 'Remove-RDPRights')]
    [Parameter(Mandatory = $true, ParameterSetName = 'Generate-RDPFiles')]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$UserListPath,

    [Parameter(Mandatory = $true, ParameterSetName = 'Set-RDPRights')]
    [Parameter(Mandatory = $true, ParameterSetName = 'Remove-RDPRights')]
    [Parameter(Mandatory = $true, ParameterSetName = 'Generate-RDPFiles')]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$ClientListPath,

    [Parameter(Mandatory = $true, ParameterSetName = 'GenerateFromLog')]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$InputLogPath,

    #--- OPTIONEN ---
    [Parameter(Mandatory = $false)]
    [Switch]$SendEmail, 

    [Parameter(Mandatory = $false)]
    [Switch]$SaveAsMsgOnly, 

    [Parameter(Mandatory = $false)]
    [string]$SmtpServer, 

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]$Credential = $null, 

    [Parameter(Mandatory = $false)]
    [string]$OutputPath, 

    #--- CSV ANPASSUNG ---
    [Parameter(Mandatory = $false)]
    [string]$UserColumn = 'sAMAccountName',

    [Parameter(Mandatory = $false)]
    [string]$ClientColumn = 'ComputerName'
)

#==============================================================================
# GLOBALE INIT
#==============================================================================
$Version = "4.0"
$GlobalErrorLog = [System.Collections.ArrayList]::new()
$GlobalTimestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$GlobalLogDir = Join-Path -Path $PSScriptRoot -ChildPath "Logs"

#==============================================================================
# HELPER FUNCTIONS
#==============================================================================

Function Get-DynamicADDomain {
    <# Ermittelt NetBIOS und FQDN der Domäne #>
    try {
        $adDomain = Get-ADDomain -ErrorAction Stop
        return [PSCustomObject]@{ NetBIOS = $adDomain.NetBIOSName; FQDN = $adDomain.DNSRoot }
    }
    catch {
        $msg = "ACHTUNG: AD-Domäne konnte nicht ermittelt werden. Nutze Fallback 'WORKGROUP'."
        Write-Warning $msg
        $GlobalErrorLog.Add($msg) | Out-Null
        return [PSCustomObject]@{ NetBIOS = "WORKGROUP"; FQDN = $null }
    }
}

Function Invoke-RemoteGroupMembership {
    <# Setzt die Gruppenmitgliedschaft via ADSI (WinNT) #>
    param($ComputerName, $UserName, $Domain, $Action, $LocalGroupName = "Remotedesktopbenutzer")

    Write-Verbose "Verbinde zu $ComputerName..."
    $status = 'Failed' 

    try {
        $group = [ADSI]"WinNT://$ComputerName/$LocalGroupName,group"
        $userPath = "WinNT://$Domain/$UserName,user"
        
        if ($Action -eq 'Add') {
            $group.Add($userPath)
            # Check
            $group.RefreshCache()
            $members = $group.Members() | ForEach-Object { $_.GetType().InvokeMember("Name", "GetProperty", $null, $_, $null) }
            if ($members -contains $UserName) { $status = 'Success' } else { $status = 'VerificationFailed' }
        }
        elseif ($Action -eq 'Remove') {
            $group.Remove($userPath)
            # Check
            $group.RefreshCache()
            $members = $group.Members() | ForEach-Object { $_.GetType().InvokeMember("Name", "GetProperty", $null, $_, $null) }
            if (-not ($members -contains $UserName)) { $status = 'Success' } else { $status = 'VerificationFailed' }
        }
    }
    catch {
        $err = $_.Exception.Message.Trim()
        if ($err -like "*bereits Mitglied*") { $status = 'AlreadyExists' }
        elseif ($err -like "*nicht Mitglied*") { $status = 'NotMember' }
        else {
            $msg = "ADSI-Fehler auf $ComputerName für $UserName: $err"
            Write-Warning $msg
            $GlobalErrorLog.Add($msg) | Out-Null
            $status = 'Failed'
        }
    }
    return $status
}

Function Create-RDPFile {
    <# Erstellt die .rdp Datei per StringBuilder #>
    param($ComputerName, $FilePath)
    
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
        $msg = "Fehler beim Schreiben von $FilePath: $($_.Exception.Message)"
        Write-Warning $msg
        $GlobalErrorLog.Add($msg) | Out-Null
        return $false
    }
}

Function Start-EmailWorkflow {
    <# Versendet oder speichert E-Mails (Outlook oder SMTP) #>
    param($UserName, $RDPFilePaths, $MSGSavePath, $Send, $SmtpServer, $Credential)
    
    # 1. Daten aus AD holen
    try {
        if (-not (Get-Module -Name ActiveDirectory)) { Import-Module ActiveDirectory -ErrorAction Stop }
        $adUser = Get-ADUser -Identity $UserName -Properties DisplayName, EmailAddress -ErrorAction Stop
    } catch {
        $msg = "AD-Fehler bei User $UserName: $($_.Exception.Message)"
        Write-Warning $msg
        if ($Global:GlobalErrorLog) { $Global:GlobalErrorLog.Add($msg) | Out-Null }
        return $false
    }
    
    $to = $adUser.EmailAddress
    $name = $adUser.DisplayName
    
    if (-not $to) {
        $msg = "Keine E-Mail-Adresse für $UserName im AD gefunden."
        Write-Warning $msg
        if ($Global:GlobalErrorLog) { $Global:GlobalErrorLog.Add($msg) | Out-Null }
        return $false
    }
    
    # 2. Body bauen
    $sb = [System.Text.StringBuilder]::new()
    $sb.AppendLine("<html><body style='font-family:Calibri, Arial, sans-serif;'>") | Out-Null
    $sb.AppendLine("<p>Hallo $name,</p><p>anbei erhalten Sie Ihre RDP-Zugriffsdateien.</p>") | Out-Null
    $sb.AppendLine("</body></html>") | Out-Null
    $body = $sb.ToString()
    $subj = "RDP-Zugang"

    # 3. Versand
    if ($SmtpServer) {
        if (-not $Send) { return $false }
        $p = @{ To=$to; From=$to; Subject=$subj; Body=$body; BodyAsHtml=$true; SmtpServer=$SmtpServer; Attachments=$RDPFilePaths; ErrorAction='Stop' }
        if ($Credential) { $p.Add('Credential', $Credential) }
        
        try {
            Send-MailMessage @p
            Write-Host " [SMTP] Gesendet an $to" -ForegroundColor Cyan
            return $true
        } catch {
            Write-Warning "SMTP-Fehler ($UserName): $($_.Exception.Message)"
            return $false
        }
    }
    else {
        # Outlook COM
        try {
            try { $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') } 
            catch { $outlook = New-Object -ComObject Outlook.Application }
            
            if (!$outlook) { throw "Outlook nicht startbar." }
            
            $mail = $outlook.CreateItem(0)
            $mail.Subject = $subj
            $mail.To = $to
            $mail.HTMLBody = $body
            foreach ($f in $RDPFilePaths) { if (Test-Path $f) { $mail.Attachments.Add($f) } }
            
            if ($MSGSavePath) {
                $mail.SaveAs($MSGSavePath, 5) # olMSG
                Write-Host " [MSG] Gespeichert: $MSGSavePath" -ForegroundColor Green
            }
            if ($Send) {
                $mail.Send()
                Write-Host " [OUTLOOK] Gesendet an $to" -ForegroundColor Cyan
            }
            return $true
        } catch {
            Write-Warning "Outlook-Fehler ($UserName): $($_.Exception.Message)"
            return $false
        } finally {
            [System.GC]::Collect()
        }
    }
}

Function Generate-SendMailsScript {
    <# Generiert das Helper-Skript für Kontext-Wechsel #>
    param($UserList, $RDPFilePaths, $OutputPath, $SmtpServer, $SaveAsMsgOnly)
    
    Write-Verbose "Erstelle 'sendMails.ps1'..."
    $resOut = Resolve-Path $OutputPath
    
    # Arrays serialisieren
    $uStr = $UserList | ForEach-Object { "'$_'" } | Join-String -Separator ","
    $rStr = $RDPFilePaths | ForEach-Object { $p=Resolve-Path $_; "'$p'" } | Join-String -Separator ","
    $smtpStr = if ($SmtpServer) { "'$SmtpServer'" } else { "`$null" }
    $sendBool = if ($SaveAsMsgOnly) { "`$false" } else { "`$true" }
    
    # Funktions-Body extrahieren für Self-Containment
    $funcDef = (Get-Command Start-EmailWorkflow).Definition

    $sb = [System.Text.StringBuilder]::new()
    $sb.AppendLine("# AUTO-GENERATED SCRIPT: $GlobalTimestamp") | Out-Null
    $sb.AppendLine("param(`$AltSmtp)") | Out-Null
    $sb.AppendLine("Function Start-EmailWorkflow { $funcDef }") | Out-Null
    $sb.AppendLine("") | Out-Null
    $sb.AppendLine("`$Users = @($uStr)") | Out-Null
    $sb.AppendLine("`$Files = @($rStr)") | Out-Null
    $sb.AppendLine("`$Smtp = $smtpStr") | Out-Null
    $sb.AppendLine("if (`$AltSmtp) { `$Smtp = `$AltSmtp }") | Out-Null
    $sb.AppendLine("`$DoSend = $sendBool") | Out-Null
    $sb.AppendLine("`$OutDir = '$resOut'") | Out-Null
    $sb.AppendLine("") | Out-Null
    $sb.AppendLine("foreach (`$u in `$Users) {") | Out-Null
    $sb.AppendLine("    `$m = Join-Path `$OutDir (`$u + '.msg')") | Out-Null
    $sb.AppendLine("    if (`$Smtp) { `$m = `$null }") | Out-Null
    $sb.AppendLine("    Write-Host 'Verarbeite' `$u ... -ForegroundColor Cyan") | Out-Null
    $sb.AppendLine("    Start-EmailWorkflow -UserName `$u -RDPFilePaths `$Files -MSGSavePath `$m -Send `$DoSend -SmtpServer `$Smtp") | Out-Null
    $sb.AppendLine("}") | Out-Null
    $sb.AppendLine("Read-Host 'Fertig.'") | Out-Null

    try {
        $f = Join-Path $OutputPath "sendMails_$($GlobalTimestamp).ps1"
        Set-Content $f -Value $sb.ToString() -Encoding UTF8
        Write-Host " [INFO] Helper-Skript erstellt: $f" -ForegroundColor Gray
    } catch {
        Write-Warning "Konnte Helper-Skript nicht erstellen."
    }
}

Function Load-CsvData {
    param($CsvPath)
    if (-not (Test-Path $CsvPath)) { Write-Error "Datei fehlt: $CsvPath"; return $null }
    try {
        $d = Import-Csv $CsvPath -Delimiter ';' -Encoding Default -ErrorAction Stop
        if (($d | Measure-Object).Count -eq 0) { Write-Error "Datei leer: $CsvPath"; return $null }
        return $d
    } catch { Write-Error "CSV-Lesefehler: $($_.Exception.Message)"; return $null }
}

Function Write-Log {
    param($Data, $Type)
    if (-not (Test-Path $GlobalLogDir)) { New-Item $GlobalLogDir -Type Directory -Force | Out-Null }
    $f = Join-Path $GlobalLogDir "Log_${Type}_$($GlobalTimestamp).csv"
    $Data | Export-Csv -Path $f -NoType -Delimiter ';' -Encoding UTF8
    Write-Host " [LOG] Protokoll gespeichert: $f" -ForegroundColor DarkGreen
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
    
    # PHASE 1: Planung
    Write-Host " [1/3] Prüfe Erreichbarkeit und plane Aktionen..." -ForegroundColor Yellow
    $Plan = @()
    
    foreach ($c in $Clients) {
        $cN = $c.$ClientCol
        if (!$cN) { continue }
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
    if ($count -eq 0) { Write-Warning "Keine Aktionen möglich."; return }
    
    # PHASE 2: Bestätigung
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

    # PHASE 3: Ausführung
    Write-Host " [3/3] Ausführung..." -ForegroundColor Cyan
    $Rep = @()
    foreach ($p in $Plan) {
        $res = Invoke-RemoteGroupMembership -ComputerName $p.Client -UserName $p.User -Domain $AD.NetBIOS -Action $p.Action
        
        # Output für den User (Feedback loop fix)
        $col = if ($res -eq 'Success') { "Green" } elseif ($res -match 'Already|NotMember') { "Gray" } else { "Red" }
        Write-Host " $res : $($p.User) -> $($p.Client)" -ForegroundColor $col
        
        $Rep += [PSCustomObject]@{ Client=$p.Client; User=$p.User; Action=$p.Action; Status=$res; Time=Get-Date }
    }
    
    Write-Log -Data $Rep -Type "Rights"
    
    # Mail-Chaining
    if ($DoMail) {
        $suc = $Rep | Where { $_.Status -in 'Success','AlreadyExists' }
        $uList = $suc | Select -Expand User -Unique
        $cList = $suc | Select -Expand Client -Unique
        if ($uList) {
            Start-FileWorkflow -UserList $uList -ClientList $cList -IsCombined $true -OutputPath $OutPath -Smtp $Smtp -Cred $Cred -UserCol "User" -ClientCol "Client"
        }
    }
}

Function Start-FileWorkflow {
    param($UserListPath, $ClientListPath, $LogPath, $FromLog, $UserList, $ClientList, $IsCombined, $OutputPath, $UserCol, $ClientCol, $SaveOnly, $Smtp, $Cred)

    # Daten laden
    $U=$null; $C=$null
    if ($IsCombined) { $U=$UserList; $C=$ClientList; $UserCol="User"; $ClientCol="Client" }
    elseif ($FromLog) {
        $l = Load-CsvData $LogPath
        if (!$l) { return }
        $suc = $l | Where { $_.Status -in 'Success','AlreadyExists' }
        $U = $suc | Select -Expand User -Unique
        $C = $suc | Select -Expand Client -Unique
        $UserCol="User"; $ClientCol="Client"
    }
    else {
        $U = Load-CsvData $UserListPath
        $C = Load-CsvData $ClientListPath
        if (!$U -or !$C) { return }
    }
    
    if (!$OutputPath) { $OutputPath = Join-Path $PSScriptRoot "RDP_Output_$GlobalTimestamp" }
    if (-not (Test-Path $OutputPath)) { New-Item $OutputPath -ItemType Directory -Force | Out-Null }
    
    # RDPs
    Write-Host "--- MODUS: Dateien erstellen ---" -ForegroundColor Cyan
    $rdps = @()
    $cObjs = if ($IsCombined -or $FromLog) { $C } else { $C.$ClientCol }
    
    foreach ($cn in $cObjs) {
        if (!$cn) { continue }
        $p = Join-Path $OutputPath "$cn.rdp"
        if (Create-RDPFile $cn $p) { $rdps += $p }
    }
    
    if ($rdps.Count -eq 0) { Write-Error "Keine RDP-Dateien."; return }
    Write-Host " $(($rdps|Measure-Object).Count) RDP-Dateien erstellt." -ForegroundColor Green
    
    # Mails
    $uObjs = if ($IsCombined -or $FromLog) { $U } else { $U.$UserCol }
    $doSend = -not $SaveOnly
    $mRep = @()
    
    Write-Host " Verarbeite E-Mails..."
    foreach ($un in $uObjs) {
        if (!$un) { continue }
        $mp = if (!$Smtp) { Join-Path $OutputPath "$un.msg" } else { $null }
        $ok = Start-EmailWorkflow -UserName $un -RDPFilePaths $rdps -MSGSavePath $mp -Send $doSend -SmtpServer $Smtp -Credential $Cred
        $mRep += [PSCustomObject]@{ User=$un; Action=(if($doSend){"Send"}else{"Save"}); Status=$ok; Time=Get-Date }
    }
    
    Write-Log -Data $mRep -Type "Mails"
    
    # Helper Script
    Generate-SendMailsScript -UserList $uObjs -RDPFilePaths $rdps -OutputPath $OutputPath -SmtpServer $Smtp -SaveAsMsgOnly $SaveOnly
    
    if ($GlobalErrorLog.Count -gt 0) {
        $ef = Join-Path $OutputPath "ERROR.TXT"
        $GlobalErrorLog | Out-File $ef
        Write-Warning "Fehler aufgetreten. Siehe: $ef"
    }
}

#==============================================================================
# MAIN
#==============================================================================
Write-Host "Manage-RDPUsers v$Version" -ForegroundColor White

if ($PSCmdlet.ParameterSetName -ne 'Interactive') {
    if ($SetRDPRights -or $RemoveRDPRights) {
        Start-RightsWorkflow -Set $SetRDPRights -UserListPath $UserListPath -ClientListPath $ClientListPath -UserCol $UserColumn -ClientCol $ClientColumn -DoMail $SendEmail -Smtp $SmtpServer -Cred $Credential -OutPath $OutputPath
    }
    elseif ($GenerateRDPFiles) {
        Start-FileWorkflow -UserListPath $UserListPath -ClientListPath $ClientListPath -OutputPath $OutputPath -UserCol $UserColumn -ClientCol $ClientColumn -SaveOnly $SaveAsMsgOnly -Smtp $SmtpServer -Cred $Credential
    }
    elseif ($GenerateFromLog) {
        Start-FileWorkflow -FromLog $true -InputLogPath $InputLogPath -OutputPath $OutputPath -SaveOnly $SaveAsMsgOnly -Smtp $SmtpServer -Cred $Credential
    }
}
else {
    # Einfaches Menü
    Write-Host " 1. Rechte setzen (Add)"
    Write-Host " 2. Rechte entfernen (Remove)"
    Write-Host " 3. Dateien erstellen (CSV)"
    Write-Host " 4. Dateien erstellen (Log)"
    $opt = Read-Host "Wahl"
    if ($opt -eq '1') {
        $u = Read-Host "User-CSV"; $c = Read-Host "Client-CSV"
        Start-RightsWorkflow -Set $true -UserListPath $u -ClientListPath $c
    }
    if ($opt -eq '3') {
        $u = Read-Host "User-CSV"; $c = Read-Host "Client-CSV"
        Start-FileWorkflow -UserListPath $u -ClientListPath $c
    }
    # (Weitere Optionen analog...)
}
