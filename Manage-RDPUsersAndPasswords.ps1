<#
.FILENAME Manage-RDPUsersAndPasswords.ps1
.DESCRIPTION
    Ein PowerShell-Skript (v5.1) zur Verwaltung von Remotedesktop-Berechtigungen 
    und zur automatisierten Erstellung von Benutzer-Anmeldeinformationen (Passwörter, RDP-Dateien, E-Mail-Entwürfe).
.NOTES
    Autor           : PS-Coding (via Gemini)
    Version         : 1.4 (Umstellung auf ADSI, Implementierung Plan/Confirm/Execute-Logik)
    Erstellt am     : 10.11.2025
    PowerShell      : 5.1
    Umgebung        : Windows On-Premise (AD)
    Voraussetzungen : 1. PowerShell 5.1
                      2. Active Directory Modul (RSAT)
                      3. Administrator-Rechte (Lokal auf Clients / Domäne für PW-Reset)
                      4. Microsoft Outlook (Desktop-Client) für Modus -ResetPasswordWorkflow
                      5. WinRM/WMI/RPC muss auf Clients erreichbar sein.

.AI-SOURCES
    - Konzept für ADSI-Gruppenverwaltung: Microsoft (TechNet/MSDN)
    - Konzept für Outlook COM-Objekt: Microsoft Developer Network (MSDN)
.USER-SOURCES
    - Keine

.LINK
    https://learn.microsoft.com/en-us/powershell/module/activedirectory/
    https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.mailitem

.PARAMETERSET Set-RDPRights
    Modus 1: Fügt Benutzer (N) zu Clients (M) hinzu. (Alle zu Allen)

.PARAMETERSET Remove-RDPRights
    Modus 3: Entfernt Benutzer (N) von Clients (M). (Alle von Allen)

.PARAMETERSET Reset-PasswordWorkflow
    Modus 2: 1:1-Workflow. Setzt Passwörter, erstellt RDP- und MSG-Dateien.
             (Benötigt Outlook)

.EXAMPLE
    # MODUS 1: Remotedesktop-Rechte HINZUFÜGEN (Alle User auf alle Clients)
    # Führt eine Planung, Bestätigung und dann die Ausführung durch.
    .\Manage-RDPUsersAndPasswords.ps1 -Set-RDPRights -UserListPath "C:\temp\userlist.csv" -ClientListPath "C:\temp\clientlist.csv" -Verbose

.EXAMPLE
    # MODUS 1 (WHATIF): Nur die Planung und Anzeige durchführen, ohne Ausführung.
    .\Manage-RDPUsersAndPasswords.ps1 -Set-RDPRights -UserListPath "C:\temp\userlist.csv" -ClientListPath "C:\temp\clientlist.csv" -WhatIf

.EXAMPLE
    # MODUS 3: Remotedesktop-Rechte ENTFERNEN (Alle User von allen Clients)
    .\Manage-RDPUsersAndPasswords.ps1 -Remove-RDPRights -UserListPath "C:\temp\userlist.csv" -ClientListPath "C:\temp\clientlist.csv" -Verbose

.EXAMPLE
    # MODUS 2: Passwort-Workflow (1:1 Zuweisung)
    .\Manage-RDPUsersAndPasswords.ps1 -ResetPasswordWorkflow -UserListPath "C:\temp\userlist.csv" -ClientListPath "C:\temp\clientlist.csv" -OutputPath "C:\temp\Ausgabe" -Verbose
#>

[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'Set-RDPRights')]
param(
    #--- MODUS 1: Schalter ---
    [Parameter(Mandatory = $true, ParameterSetName = 'Set-RDPRights')]
    [Switch]
    $SetRDPRights,

    #--- MODUS 3: Schalter ---
    [Parameter(Mandatory = $true, ParameterSetName = 'Remove-RDPRights')]
    [Switch]
    $RemoveRDPRights,

    #--- MODUS 2: Schalter ---
    [Parameter(Mandatory = $true, ParameterSetName = 'ResetPasswordWorkflow')]
    [Switch]
    $ResetPasswordWorkflow,

    #--- Gemeinsame Parameter für alle Modi ---
    [Parameter(Mandatory = $true, ParameterSetName = 'Set-RDPRights')]
    [Parameter(Mandatory = $true, ParameterSetName = 'Remove-RDPRights')]
    [Parameter(Mandatory = $true, ParameterSetName = 'ResetPasswordWorkflow')]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]
    $UserListPath,

    [Parameter(Mandatory = $true, ParameterSetName = 'Set-RDPRights')]
    [Parameter(Mandatory = $true, ParameterSetName = 'Remove-RDPRights')]
    [Parameter(Mandatory = $true, ParameterSetName = 'ResetPasswordWorkflow')]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]
    $ClientListPath,

    #--- Parameter für Modus 2 (Passwort-Workflow) ---
    [Parameter(Mandatory = $true, ParameterSetName = 'ResetPasswordWorkflow')]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]
    $OutputPath,

    #--- Optionale Parameter (CSV-Header) ---
    [Parameter(Mandatory = $false)]
    [string]
    $UserColumn = 'sAMAccountName',

    [Parameter(Mandatory = $false)]
    [string]
    $ClientColumn = 'ComputerName',
    
    #--- Optionale Parameter (Passwort-Generierung) ---
    [Parameter(Mandatory = $false, ParameterSetName = 'ResetPasswordWorkflow')]
    [int]
    $PasswordLength = 10
)

#==============================================================================
# HILFSFUNKTIONEN
#==============================================================================

Function Get-DynamicADDomain {
    <#
    .SYNOPSIS
        Ermittelt dynamisch den AD-Domänenkontext (NetBIOS und FQDN).
    #>
    try {
        $adDomain = Get-ADDomain -ErrorAction Stop
        $netBIOS = $adDomain.NetBIOSName
        $fqdn = $adDomain.DNSRoot
        
        Write-Verbose "AD-Domäne dynamisch erkannt: $netBIOS ($fqdn)"
        return [PSCustomObject]@{
            NetBIOS = $netBIOS
            FQDN    = $fqdn
        }
    }
    catch {
        Write-Warning "Fehler beim dynamischen Ermitteln der AD-Domäne. Stelle sicher, dass das AD-Modul geladen ist und eine Domänenverbindung besteht."
        Write-Warning "Verwende Fallback 'WORKGROUP' für lokale Konten (kann fehlschlagen, wenn Domänenkonten erwartet werden)."
        return [PSCustomObject]@{
            NetBIOS = "WORKGROUP" # Fallback für lokale Gruppen
            FQDN    = $null
        }
    }
}

Function Invoke-RemoteGroupMembership {
    <#
    .SYNOPSIS
        (Version 1.4) Fügt einen Benutzer zu einer lokalen Gruppe auf einem Remote-Client hinzu oder entfernt ihn.
        Nutzt [ADSI] (WinNT) für maximale Kompatibilität und Robustheit (PS 5.1).
    #>
    param(
        [string]$ComputerName,
        [string]$UserName,
        [string]$Domain,
        [string]$Action, # 'Add' or 'Remove'
        [string]$LocalGroupName = "Remotedesktopbenutzer"
    )

    Write-Verbose "Aktion '$Action': User '$Domain\$UserName' -> Gruppe '$LocalGroupName' auf '$ComputerName'"

    try {
        # 1. Verbindung zur LOKALEN Gruppe auf dem ZIEL-PC herstellen
        $group = [ADSI]("WinNT://$ComputerName/$LocalGroupName,group")
        
        # 2. Verbindung zum DOMÄNEN-Benutzer herstellen (aus Sicht des Admin-PCs)
        # ADSI benötigt den Pfad zum Benutzerobjekt
        $userPath = "WinNT://$Domain/$UserName,user"
        
        # 3. Aktion ausführen
        if ($Action -eq 'Add') {
            Write-Verbose "Füge hinzu: $userPath ZU $($group.Path)"
            $group.Add($userPath)
            Write-Host " ERFOLG: $Domain\$UserName zu $LocalGroupName auf $ComputerName HINZUGEFÜGT." -ForegroundColor Green
        }
        elseif ($Action -eq 'Remove') {
            Write-Verbose "Entferne: $userPath VON $($group.Path)"
            $group.Remove($userPath)
            Write-Host " ERFOLG: $Domain\$UserName von $LocalGroupName auf $ComputerName ENTFERNT." -ForegroundColor Yellow
        }
        
        # 4. Erfolg prüfen (ADSI wirft bei Fehlern Exceptions, die im Catch landen)
        # Manuelle Prüfung (IsMember) ist mit WinNT-Provider langsam, daher verlassen wir uns auf den try/catch.
        Write-Verbose "Aktion '$Action' für $UserName auf $ComputerName war erfolgreich."
        return $true
    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Warning ("ADSI-FEHLER bei Client {0} für User {1}: {2}" -f $ComputerName, $UserName, $errorMessage.Trim())
        Write-Warning "Mögliche Ursachen: Admin-Rechte fehlen, RPC/Netzwerk-Firewall, Benutzer/Gruppe nicht gefunden, Benutzer bereits Mitglied (beim Hinzufügen) oder kein Mitglied (beim Entfernen)."
        return $false
    }
}

Function New-RandomStrongPassword {
    <#
    .SYNOPSIS
        Generiert ein starkes, aber relativ merkbares Passwort (Wort-Sonderzeichen-Zahl-Wort).
    #>
    param (
        [int]$MinLength = 8,
        [int]$MaxLength = 10
    )

    # Einfache, deutsche Wörter (können erweitert werden)
    $words = @(
        "Sonne", "Mond", "Stern", "Haus", "Auto", "Tisch", "Stuhl", "Blau", "Rot", "Grün",
        "Apfel", "Birne", "Reise", "Weg", "Ziel", "Test", "Daten", "Code", "Netz", "Fluss"
    )
    
    # Sonderzeichen (ohne Verwechslungsgefahr, z.B. kein O/0, l/1)
    $chars = @('!', '?', '#', '$', '%', '&', '*', ':', '-', '_')
    
    # Zufällige Auswahl
    $word1 = $words | Get-Random
    $word2 = $words | Get-Random
    $char = $chars | Get-Random
    $num = Get-Random -Minimum 100 -Maximum 9999
    
    # Muster-Zusammenstellung (zufällig)
    $pattern = Get-Random -Minimum 1 -Maximum 4
    $password = ""

    switch ($pattern) {
        1 { $password = "$($word1)$($char)$($num)$($word2)" }
        2 { $password = "$($num)$($char)$($word1)$($word2)" }
        3 { $password = "$($word1)$($word2)$($char)$($num)" }
        4 { $password = "$($word1)$($num)$($char)$($word2)" }
    }
    
    # Längenprüfung (falls die Wörter zu lang/kurz waren)
    if ($password.Length -gt $MaxLength) {
        $password = $password.Substring(0, $MaxLength)
    }
    if ($password.Length -lt $MinLength) {
        # Falls zu kurz (unwahrscheinlich), fülle mit Zahlen auf
        $password = $password.PadRight($MinLength, (Get-Random -Minimum 0 -Maximum 9))
    }
    
    Write-Verbose "Generiertes Passwort: $password (Länge: $($password.Length))"
    return $password
}

Function Create-RDPFile {
    <#
    .SYNOPSIS
        Erstellt eine .RDP-Konfigurationsdatei.
    #>
    param(
        [string]$ComputerName,
        [string]$UserName,
        [string]$Domain,
        [string]$FilePath
    )

    Write-Verbose "Erstelle RDP-Datei für $UserName@$ComputerName in '$FilePath'"
    
    # Standard-RDP-Datei-Inhalt
    $rdpContent = @"
screen mode id:i:2
desktopwidth:i:1920
desktopheight:i:1080
session bpp:i:32
winposstr:s:0,1,100,100,800,600
compression:i:1
keyboardhook:i:2
displayconnectionbar:i:1
disable wallpaper:i:1
allow font smoothing:i:1
allow desktop composition:i:1
disable full window drag:i:1
disable menu anims:i:1
disable themes:i:1
disable cursor setting:i:0
bitmapcachepersistenable:i:1
full address:s:$ComputerName
audiomode:i:0
redirectdrives:i:0
redirectprinters:i:0
redirectcomports:i:0
redirectsmartcards:i:0
drivestoredirect:s:
redirectclipboard:i:1
redirectposdevices:i:0
authentication level:i:2
prompt for credentials:i:1
negotiate security layer:i:1
remoteapplicationmode:i:0
use multimon:i:0
enablerdsa:i:0
username:s:$Domain\$UserName
domain:s:$Domain
disable themes:i:1
connect to console:i:0
gatewayusagemethod:i:2
promptcredentialonce:i:1
"@

    try {
        Set-Content -Path $FilePath -Value $rdpContent -Encoding UTF8 -ErrorAction Stop
        if (Test-Path $FilePath) {
            Write-Host " ERFOLG: RDP-Datei erstellt: $FilePath" -ForegroundColor Green
            return $true
        }
    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Warning ("Fehler beim Schreiben der RDP-Datei '{0}': {1}" -f $FilePath, $errorMessage)
        return $false
    }
}

Function Create-OutlookMSG {
    <#
    .SYNOPSIS
        Erstellt eine .MSG E-Mail-Datei mit Anhängen. (Version 1.2)
        Voraussetzung: Outlook Desktop-Client muss installiert sein.
    #>
    param(
        [string]$RecipientName, # z.B. Max Mustermann
        [string]$UserName,
        [string]$Password,
        [string]$ComputerName,
        [string]$RDPFilePath, # Pfad zur .rdp Datei für den Anhang
        [string]$MSGSavePath  # Ziel-Pfad für die .msg Datei
    )

    Write-Verbose "Erstelle Outlook .MSG-Datei für $UserName..."
    
    $Error.Clear()
    $outlook = $null
    
    try {
        # Prüfen, ob Outlook läuft, sonst starten
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
        
        # 0 = olMailItem
        $mail = $outlook.CreateItem(0) 

        # E-Mail-Inhalt (Anpassbar)
        $mail.Subject = "Ihre neuen RDP-Zugangsdaten"
        $mail.HTMLBody = @"
<html>
<body style="font-family:Calibri, Arial, sans-serif; font-size:11pt;">
<p>Hallo $RecipientName,</p>
<p>hier sind Ihre persönlichen Zugangsdaten für den Remote-Desktop-Zugriff.</p>
<p>
    <b>Client (Computer):</b> $ComputerName<br>
    <b>Benutzername:</b> $UserName<br>
    <b>Kennwort:</b> <span style="font-family:Consolas, monospace; background-color:#f0f0f0; padding: 2px 4px; border: 1px solid #ccc;">$Password</span>
</p>
<p>
    <b>Anleitung:</b><br>
    1. Speichern Sie die angehängte <b>.rdp-Datei</b> auf Ihrem Desktop.<br>
    2. Öffnen Sie die Datei per Doppelklick.<br>
    3. Geben Sie das oben genannte Kennwort ein, wenn Sie dazu aufgefordert werden.
</p>
<p>
    <i>(Dies ist eine automatisch generierte Nachricht.)</i>
</p>
</body>
</html>
"@

        # Anhang hinzufügen
        if (Test-Path $RDPFilePath) {
            $attachment = $mail.Attachments.Add($RDPFilePath, 1, 1, $RDPFilePath.Split('\')[-1])
            Write-Verbose "Anhang '$RDPFilePath' hinzugefügt."
        } else {
            Write-Warning "RDP-Anhang '$RDPFilePath' nicht gefunden. E-Mail wird ohne Anhang erstellt."
        }

        # 5 = olMSG (Format)
        $mail.SaveAs($MSGSavePath, 5) 
        
        Write-Host " ERFOLG: MSG-Datei gespeichert: $MSGSavePath" -ForegroundColor Green
        return $true
    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Warning ("Fehler beim Erstellen der MSG-Datei für {0}: {1}" -f $UserName, $errorMessage)
        
        if ($Error) {
            Write-Warning "Letzter COM-Fehler: $($Error[0])"
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

#==============================================================================
# DATEN LADEN
#==============================================================================

Write-Verbose "Lade CSV-Dateien..."
$CsvImportParams = @{
    Delimiter = ';'
    Encoding  = 'Default' # Sicherer für deutsche Excel-Exporte
    ErrorAction = 'Stop'
}

try {
    $Users = Import-Csv -Path $UserListPath @CsvImportParams
    $Clients = Import-Csv -Path $ClientListPath @CsvImportParams
}
catch {
    Write-Error "Fehler beim Lesen der CSV-Dateien. Stelle sicher, dass das Trennzeichen ';' ist und die Dateien existieren."
    $errorMessage = $_.Exception.Message
    Write-Error ("Details: {0}" -f $errorMessage)
    return
}

Write-Host "Daten geladen: $($Users.Count) Benutzer und $($Clients.Count) Clients gefunden."

#==============================================================================
# MODUS-AUSFÜHRUNG
#==============================================================================

# --- Dynamische AD-Daten holen ---
if (-not (Get-Module -Name ActiveDirectory)) {
    try {
        Write-Verbose "Versuche, das ActiveDirectory-Modul zu importieren..."
        Import-Module ActiveDirectory -ErrorAction Stop
    } catch {
        Write-Warning "Das PowerShell-Modul 'ActiveDirectory' konnte nicht geladen werden."
        Write-Warning "Für Modus 2 (Passwort) und Modus 1/3 (Domänenkonten) ist dies erforderlich."
        if ($ResetPasswordWorkflow) {
            Write-Error "Modus 2 kann ohne das AD-Modul nicht ausgeführt werden. Skript wird beendet."
            return
        }
    }
}
$ADContext = Get-DynamicADDomain


# --- MODUS 1 & 3: BERECHTIGUNGEN (N:M) ---
if ($SetRDPRights -or $RemoveRDPRights) {
    
    $CurrentAction = if ($SetRDPRights) { 'Add' } else { 'Remove' }
    $ActionVerb = if ($SetRDPRights) { "Hinzufügen" } else { "Entfernen" }
    
    Write-Host "Starte Modus: Remotedesktop-Berechtigungen ($ActionVerb) (Alle User -> Alle Clients)" -ForegroundColor Cyan
    Write-Host "Betrifft $($Users.Count) Benutzer auf $($Clients.Count) Clients."

    # --- PHASE 1: PLANUNG & VORAB-TESTS (NEU V1.4) ---
    Write-Host "PHASE 1: Plane Aktionen und teste Erreichbarkeit (Ping)..." -ForegroundColor Yellow
    $plannedActions = @()
    $offlineClients = @()

    foreach ($client in $Clients) {
        $clientName = $client.$($ClientColumn)
        if ([string]::IsNullOrWhiteSpace($clientName)) {
            Write-Warning "Eintrag in Client-Liste übersprungen (Name ist leer)."
            continue
        }
        
        Write-Verbose "Teste Erreichbarkeit von $clientName..."
        if (-not (Test-Connection -ComputerName $clientName -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
            Write-Warning "Client $clientName ist offline (Ping fehlgeschlagen). Alle Aktionen für diesen Client werden übersprungen."
            $offlineClients += $clientName
            continue
        }
        
        Write-Verbose "Client $clientName ist online. Plane Aktionen..."
        
        # Client ist online, plane Aktionen für alle User
        foreach ($user in $Users) {
            $userName = $user.$($UserColumn)
            if ([string]::IsNullOrWhiteSpace($userName)) {
                Write-Warning "Eintrag in Benutzer-Liste übersprungen (Name ist leer)."
                continue
            }
            
            $plannedActions += [PSCustomObject]@{
                Client     = $clientName
                User       = $userName
                Action     = $CurrentAction
            }
        }
    }
    
    # --- PHASE 2: ANZEIGE & BESTÄTIGUNG (NEU V1.4) ---
    Write-Host "-------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host "PHASE 2: Planungsübersicht & Bestätigung" -ForegroundColor Cyan
    Write-Host "-------------------------------------------------------------"
    
    if ($offlineClients.Count -gt 0) {
        Write-Host "Die folgenden $($offlineClients.Count) Clients sind OFFLINE und werden übersprungen:" -ForegroundColor Yellow
        $offlineClients | ForEach-Object { Write-Host " - $_" }
    }
    
    if ($plannedActions.Count -eq 0) {
        Write-Host "Keine Aktionen für Online-Clients geplant." -ForegroundColor Green
        Write-Host "Skript-Ausführung beendet."
        return
    }

    Write-Host "Es sind $($plannedActions.Count) Aktionen für Online-Clients geplant:" -ForegroundColor Green
    $plannedActions | Format-Table -AutoSize
    
    Write-Host "-------------------------------------------------------------"
    
    # -WhatIf Support: Zeigt den Plan, aber fragt nicht nach Ausführung.
    if ($pscmdlet.ShouldProcess("die $($plannedActions.Count) oben gelisteten Aktionen", "Ausführen")) {
        
        # Interaktive Bestätigung, WENN -WhatIf NICHT verwendet wird
        $confirmation = Read-Host -Prompt "Möchten Sie diese $($plannedActions.Count) Aktionen jetzt ausführen? (J/N)"
        if ($confirmation -ne 'J') {
            Write-Warning "Aktion vom Benutzer abgebrochen."
            return
        }

        # --- PHASE 3: AUSFÜHRUNG (NEU V1.4) ---
        Write-Host "-------------------------------------------------------------" -ForegroundColor Cyan
        Write-Host "PHASE 3: Führe Aktionen aus..." -ForegroundColor Cyan
        Write-Host "-------------------------------------------------------------"
        
        $Report = @()
        foreach ($action in $plannedActions) {
            
            $success = Invoke-RemoteGroupMembership -ComputerName $action.Client -UserName $action.User -Domain $ADContext.NetBIOS -Action $action.Action
            
            $Report += [PSCustomObject]@{
                Client     = $action.Client
                User       = $action.User
                Action     = $action.Action
                Success    = $success
                Timestamp  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
        
        Write-Host "Modus ($ActionVerb) abgeschlossen." -ForegroundColor Cyan
        $Report | Format-Table
    }
    else {
        Write-Host "Aktion übersprungen (-WhatIf Parameter wurde verwendet)." -ForegroundColor Yellow
    }
}

# --- MODUS 2: PASSWORT-WORKFLOW (1:1) ---
if ($ResetPasswordWorkflow) {
    
    Write-Host "Starte Modus: Passwort-Workflow (1:1 Zuweisung)" -ForegroundColor Cyan
    
    # 1. Validierung der Listenlänge
    if ($Users.Count -gt $Clients.Count) {
        Write-Error "Fehler: Es sind mehr Benutzer ($($Users.Count)) als Clients ($($Clients.Count)) vorhanden."
        Write-Error "Eine 1:1-Zuweisung ist nicht möglich. Skript wird beendet."
        return
    }
    if ($Users.Count -lt $Clients.Count) {
        Write-Warning "Hinweis: Es sind weniger Benutzer ($($Users.Count)) als Clients ($($Clients.Count)) vorhanden."
        Write-Warning "Überzählige Clients werden ignoriert."
    }
    
    Write-Host "Listen-Validierung bestanden. $($Users.Count) Paarungen werden verarbeitet."
    
    # --- PHASE 1 (Modus 2): PLANUNG & VORAB-TESTS ---
    Write-Host "PHASE 1: Plane Aktionen und teste Erreichbarkeit (Ping)..." -ForegroundColor Yellow
    $plannedPairs = @()
    $offlinePairs = @()
    
    for ($i = 0; $i -lt $Users.Count; $i++) {
        $currentUser = $Users[$i]
        $currentClient = $Clients[$i]
        
        $userName = $currentUser.$($UserColumn)
        $clientName = $currentClient.$($ClientColumn)
        
        if ([string]::IsNullOrWhiteSpace($userName) -or ([string]::IsNullOrWhiteSpace($clientName))) {
            Write-Warning "Paar übersprungen (Index $i: Benutzer oder Client-Name ist leer)."
            continue
        }

        $pairInfo = [PSCustomObject]@{
            Index      = $i
            User       = $userName
            Client     = $clientName
            ADUser     = $null # Wird später gefüllt
            ADClient   = $null # Wird später gefüllt
        }

        Write-Verbose "Teste Erreichbarkeit von $clientName..."
        if (-not (Test-Connection -ComputerName $clientName -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
            Write-Warning "Client $clientName ist offline (Ping fehlgeschlagen). Paarung für $userName wird übersprungen."
            $offlinePairs += $pairInfo
            continue
        }
        
        $plannedPairs += $pairInfo
    }
    
    # --- PHASE 2 (Modus 2): ANZEIGE & BESTÄTIGUNG ---
    Write-Host "-------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host "PHASE 2: Planungsübersicht & Bestätigung" -ForegroundColor Cyan
    Write-Host "-------------------------------------------------------------"
    
    if ($offlinePairs.Count -gt 0) {
        Write-Host "Die folgenden $($offlinePairs.Count) Paare werden übersprungen (Client ist OFFLINE):" -ForegroundColor Yellow
        $offlinePairs | Format-Table -AutoSize
    }
    
    if ($plannedPairs.Count -eq 0) {
        Write-Host "Keine Aktionen für Online-Paare geplant." -ForegroundColor Green
        Write-Host "Skript-Ausführung beendet."
        return
    }

    Write-Host "Es sind $($plannedPairs.Count) Aktionen (Passwort-Reset, RDP, MSG) geplant:" -ForegroundColor Green
    $plannedPairs | Format-Table -AutoSize
    
    Write-Host "-------------------------------------------------------------"

    # -WhatIf Support
    if ($pscmdlet.ShouldProcess("die $($plannedPairs.Count) oben gelisteten Passwort-Workflows", "Ausführen")) {
        
        $confirmation = Read-Host -Prompt "Möchten Sie diese $($plannedPairs.Count) Workflows jetzt ausführen? (J/N)"
        if ($confirmation -ne 'J') {
            Write-Warning "Aktion vom Benutzer abgebrochen."
            return
        }

        # --- PHASE 3 (Modus 2): AUSFÜHRUNG ---
        Write-Host "-------------------------------------------------------------" -ForegroundColor Cyan
        Write-Host "PHASE 3: Führe Workflows aus..." -ForegroundColor Cyan
        Write-Host "-------------------------------------------------------------"
        
        # Ausgabeordner sicherstellen
        if (-not (Test-Path $OutputPath)) {
            Write-Verbose "Erstelle Ausgabeordner: $OutputPath"
            New-Item -Path $OutputPath -ItemType Directory | Out-Null
        }
        
        $PasswordReport = @()
        
        foreach ($pair in $plannedPairs) {
            
            $userName = $pair.User
            $clientName = $pair.Client
            
            Write-Host "--- Verarbeite Paar: $userName <-> $clientName ---" -ForegroundColor Yellow
            
            # 3. Schritt: Passwort generieren
            $newPassword = New-RandomStrongPassword -MinLength 8 -MaxLength $PasswordLength
            
            # 4. Schritt: Passwort im AD setzen
            try {
                Write-Verbose "Setze AD-Passwort für '$userName'..."
                $adUser = Get-ADUser -Identity $userName -ErrorAction Stop
                
                $pwSecure = $newPassword | ConvertTo-SecureString -AsPlainText -Force
                Set-ADAccountPassword -Identity $adUser -NewPassword $pwSecure -Reset -ErrorAction Stop
                Set-ADUser -Identity $adUser -ChangePasswordAtLogon $false -ErrorAction Stop
                
                Write-Host " ERFOLG: AD-Passwort für $userName zurückgesetzt." -ForegroundColor Green
                
                $PasswordReport += [PSCustomObject]@{
                    sAMAccountName = $userName
                    NeuesKennwort  = $newPassword
                    Client         = $clientName
                }
            }
            catch {
                $errorMessage = $_.Exception.Message
                Write-Warning ("Fehler beim Setzen des AD-Passworts für {0}: {1}" -f $userName, $errorMessage)
                Write-Warning "Dieser Benutzer wird für RDP/MSG-Erstellung übersprungen."
                continue # Nächster Benutzer
            }
            
            # 5. Schritt: RDP-Datei erstellen
            $safeUserName = $userName -replace '[\\/:*?"<>|]', '_'
            $rdpFileName = "$($safeUserName)_$($clientName).rdp"
            $rdpFullPath = Join-Path -Path $OutputPath -ChildPath $rdpFileName
            
            Create-RDPFile -ComputerName $clientName -UserName $userName -Domain $ADContext.NetBIOS -FilePath $rdpFullPath
            
            # 6. Schritt: MSG-Datei erstellen
            $msgFileName = "$($safeUserName)_$($clientName).msg"
            $msgFullPath = Join-Path -Path $OutputPath -ChildPath $msgFileName
            
            $displayName = try { (Get-ADUser $userName -Properties DisplayName).DisplayName } catch { $userName }

            Create-OutlookMSG -RecipientName $displayName -UserName $userName -Password $newPassword -ComputerName $clientName -RDPFilePath $rdpFullPath -MSGSavePath $msgFullPath
        }
        
        # 7. Schritt: Passwort-Export
        if ($PasswordReport.Count -gt 0) {
            $pwExportPath = Join-Path -Path $OutputPath -ChildPath "Passwort_Export.csv"
            Write-Host "Exportiere $($PasswordReport.Count) neue Passwörter nach: $pwExportPath"
            $PasswordReport | Export-Csv -Path $pwExportPath -NoTypeInformation -Delimiter ';' -Encoding UTF8
        }
    }
    else {
        Write-Host "Aktion übersprungen (-WhatIf Parameter wurde verwendet)." -ForegroundColor Yellow
    }
    
    Write-Host "Modus (Passwort-Workflow) abgeschlossen." -ForegroundColor Cyan
    Write-Host "Alle Ausgaben finden Sie in: $OutputPath"
}

Write-Host "Skript-Ausführung beendet."
