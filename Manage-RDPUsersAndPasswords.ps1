<#
.FILENAME Manage-RDPUsers.ps1
.DESCRIPTION
    Ein PowerShell-Skript (v5.1) zur Verwaltung von lokalen Remotedesktop-Berechtigungen auf Clients
    und zur automatisierten Erstellung und Verteilung von RDP-Dateien (inkl. E-Mail-Workflow).
    
    Zweck:
    - Massenhafte Zuweisung/Entfernung von Usern zu Clients (N:M).
    - Generierung personalisierter RDP-Dateien.
    - Verteilung per E-Mail (Outlook oder SMTP) oder Bereitstellung als .MSG-Dateien.

.NOTES
    Version         : 4.1
    Autor           : PS-Coding
    Erstellt am     : 18.11.2025
    PowerShell      : 5.1
    Umgebung        : Windows On-Premise (AD)

.AI-SOURCES
    - Microsoft TechNet: ADSI/WinNT Provider für lokale Gruppenverwaltung
    - MSDN: Outlook.Application COM Objekt Referenz

.USER-SOURCES
    - Anforderungen: 1:1 Zuweisung entfernt, N:M Logik implementiert.
    - Anforderungen: SMTP und Outlook Unterstützung.
    - Anforderungen: Generierung eines Helper-Skripts für Kontextwechsel.

.LINK
    https://learn.microsoft.com/en-us/powershell/module/activedirectory/

.PARAMETER SetRDPRights
    Modus 1: Fügt Benutzer (aus CSV) zur lokalen Gruppe 'Remotedesktopbenutzer' auf Clients (aus CSV) hinzu.

.PARAMETER RemoveRDPRights
    Modus 2: Entfernt Benutzer (aus CSV) aus der lokalen Gruppe 'Remotedesktopbenutzer' auf Clients (aus CSV).

.PARAMETER GenerateRDPFiles
    Modus 3: Erstellt RDP-Dateien für alle Clients und sendet diese per E-Mail an alle Benutzer.
    Standard: Direkter Versand.
    Alias: generatemail

.PARAMETER GenerateFromLog
    Modus 4: Liest ein Log-File eines vorherigen Durchlaufs und startet den E-Mail-Workflow für erfolgreiche Einträge.
    Alias: genfromlog

.PARAMETER UserListPath
    Pfad zur CSV-Datei mit Benutzernamen. Erwartet Header: sAMAccountName (konfigurierbar).

.PARAMETER ClientListPath
    Pfad zur CSV-Datei mit Computernamen. Erwartet Header: ComputerName (konfigurierbar).

.PARAMETER InputLogPath
    Pfad zur CSV-Log-Datei, die als Eingabe für Modus 4 dient.

.PARAMETER SendEmail
    Optionaler Schalter für Modus 1 & 2: Startet nach der Rechtevergabe automatisch den E-Mail-Workflow.

.PARAMETER SaveAsMsgOnly
    Schalter für Modus 3 & 4: Speichert E-Mails nur als .MSG-Entwürfe, sendet sie aber NICHT.

.PARAMETER SmtpServer
    Optional: Adresse eines SMTP-Servers. Wenn angegeben, wird SMTP statt Outlook verwendet (erfordert keine lokale Outlook-Installation).

.PARAMETER Credential
    Optional: Anmeldeinformationen für den SMTP-Server.

.EXAMPLE
    # 1. Rechte setzen (Interaktive Bestätigung):
    .\Manage-RDPUsers.ps1 -SetRDPRights -UserListPath "users.csv" -ClientListPath "clients.csv"

.EXAMPLE
    # 2. Rechte setzen UND E-Mails via SMTP senden (Automatisierung):
    $cred = Get-Credential
    .\Manage-RDPUsers.ps1 -SetRDPRights -UserListPath "users.csv" -ClientListPath "clients.csv" -SendEmail -SmtpServer "smtp.intern" -Credential $cred

.EXAMPLE
    # 3. Nur Dateien erstellen (Outlook Modus):
    .\Manage-RDPUsers.ps1 -GenerateRDPFiles -UserListPath "users.csv" -ClientListPath "clients.csv" -Verbose
#>

[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'Interactive')]
param(
    #--- MODUS-SCHALTER (ParameterSetNames nun konsistent ohne Bindestriche im Namen) ---
    [Parameter(Mandatory = $false, ParameterSetName = 'SetRDPRights')]
    [Switch]$SetRDPRights,

    [Parameter(Mandatory = $false, ParameterSetName = 'RemoveRDPRights')]
    [Switch]$RemoveRDPRights,

    [Parameter(Mandatory = $false, ParameterSetName = 'GenerateRDPFiles')]
    [Alias('generatemail')]
    [Switch]$GenerateRDPFiles,

    [Parameter(Mandatory = $false, ParameterSetName = 'GenerateFromLog')]
    [Alias('genfromlog')]
    [Switch]$GenerateFromLog,
    
    #--- EINGABE-DATEIEN ---
    [Parameter(Mandatory = $true, ParameterSetName = 'SetRDPRights')]
    [Parameter(Mandatory = $true, ParameterSetName = 'RemoveRDPRights')]
    [Parameter(Mandatory = $true, ParameterSetName = 'GenerateRDPFiles')]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$UserListPath,

    [Parameter(Mandatory = $true, ParameterSetName = 'SetRDPRights')]
    [Parameter(Mandatory = $true, ParameterSetName = 'RemoveRDPRights')]
    [Parameter(Mandatory = $true, ParameterSetName = 'GenerateRDPFiles')]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$ClientListPath,

    [Parameter(Mandatory = $true, ParameterSetName = 'GenerateFromLog')]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$InputLogPath,

    #--- GLOBALE OPTIONEN ---
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

    #--- CSV KONFIGURATION ---
    [Parameter(Mandatory = $false)]
    [string]$UserColumn = 'sAMAccountName',

    [Parameter(Mandatory = $false)]
    [string]$ClientColumn = 'ComputerName'
)

#==============================================================================
# VARIABLEN-DEFINITION
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
        return [PSCustomObject]@{ NetBIOS = $adDomain.NetBIOSName; FQDN = $adDomain.DNSRoot }
    }
    catch {
        $msg = "ACHTUNG: AD-Domäne konnte nicht ermittelt werden (AD-Modul Fehler). Nutze Fallback 'WORKGROUP'."
        Write-Warning $msg
        $GlobalErrorLog.Add($msg) | Out-Null
        return [PSCustomObject]@{ NetBIOS = "WORKGROUP"; FQDN = $null }
    }
}

Function Invoke-RemoteGroupMembership {
    param($ComputerName, $UserName, $Domain, $Action, $LocalGroupName = "Remotedesktopbenutzer")

    Write-Verbose "Verbinde zu $ComputerName..."
    $status = 'Failed' 

    try {
        $groupPath = "WinNT://$ComputerName/$LocalGroupName,group"
        $group = [ADSI]$groupPath
        $userPath = "WinNT://$Domain/$UserName,user"
        
        if ($Action -eq 'Add') {
            $group.Add($userPath)
            # Erfolg prüfen
            $group.RefreshCache()
            $members = $group.Members() | ForEach-Object { $_.GetType().InvokeMember("Name", "GetProperty", $null, $_, $null) }
            if ($members -contains $UserName) { $status = 'Success' } else { $status = 'VerificationFailed' }
        }
        elseif ($Action -eq 'Remove') {
            $group.Remove($userPath)
            # Erfolg prüfen
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
    param($ComputerName, $FilePath)
    
    # StringBuilder für Robustheit
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
    
    # FIX Zeile 252: Einfacher Check statt [string]::IsNullOrWhiteSpace
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
    $sb.AppendLine("<p><i>Automatisch generiert</i></p>") | Out-Null
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
            
            # FIX Zeile 292: Einfacher Check statt [string]::IsNullOrWhiteSpace
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
    <# Erstellt das Helper-Skript OHNE Here-Strings um Parser-Fehler zu vermeiden #>
    param($UserList, $RDPFilePaths, $OutputPath, $SmtpServer, $SaveAsMsgOnly)
    
    Write-Verbose "Erstelle 'sendMails.ps1'..."
    $resOut = Resolve-Path $OutputPath
    
    $uStr = $UserList | ForEach-Object { "'$_'" } | Join-String -Separator ","
    $rStr = $RDPFilePaths | ForEach-Object { $p=Resolve-Path $_; "'$p'" } | Join-String -Separator ","
    $smtpStr = if ($SmtpServer) { "'$SmtpServer'" } else { "`$null" }
    $sendBool = if ($SaveAsMsgOnly) { "`$false" } else { "`$true" }
    
    # Funktions-Body holen
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
    
    # Planung
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
    
    if (-not (Get-Module -Name ActiveDirectory)) { Import-Module ActiveDirectory }
    
    if (!$OutputPath) { $OutputPath = Join-Path $PSScriptRoot "RDP_Output_$GlobalTimestamp" }
    if (-not (Test-Path $OutputPath)) { New-Item $OutputPath -ItemType Directory -Force | Out-Null }
    
    Write-Host "--- MODUS: Dateien erstellen ---" -ForegroundColor Cyan
    
    # RDPs
    $rdps = @()
    $cObjs = if ($IsCombined -or $FromLog) { $C } else { $C.$ClientCol }
    
    foreach ($cn in $cObjs) {
        if (!$cn) { continue }
        if (-not (Test-Connection $cn -Count 1 -Quiet -ErrorAction SilentlyContinue)) { Write-Warning "$cn offline"; continue }
        $p = Join-Path $OutputPath "$cn.rdp"
        if (Create-RDPFile $cn $p) { $rdps += $p }
    }
    
    if ($rdps.Count -eq 0) { Write-Error "Keine RDPs erstellt."; return }
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
    # Einfaches Menü
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
    # Weitere Logik für Menü hier möglich...
    if ($sel -eq 'Q') { return }
}
Write-Host "Ende."
