<#
.SYNOPSIS
    Führt eine vollständige SCCM-Synchronisierung (Server + Client) für eine Liste von
    Computern aus, nachdem der Benutzer die auslösende(n) AD-Gruppe(n)
    aus einer interaktiven Liste ausgewählt hat.

.DESCRIPTION
    Dieses Skript ist ein All-in-One-Tool zur Beschleunigung von SCCM-Bereitstellungen,
    insbesondere in Umgebungen mit Identitätsmanagern (wie Dell One Identity).

    STANDARDVERHALTEN (FullSync):
    1. Nimmt eine Liste von Computernamen (-ComputerName) entgegen.
    2. Ermittelt den SCCM SiteCode und SiteServer FQDN dynamisch.
    3. Durchsucht SCCM nach allen AD-Gruppen, die in Sammlungen mit Bereitstellungen
       verwendet werden.
    4. Zeigt diese Gruppen in einem interaktiven GUI-Fenster (Out-GridView) zur Auswahl an.
    5. Löst die "Active Directory Group Discovery" auf dem Site Server aus.
    6. Löst die "Collection Membership Evaluation" für die ausgewählten Gruppen aus.
    7. Löst alle wichtigen Client-Aktionszyklen auf JEDEM der angegebenen Computer aus.

    ALTERNATIVES VERHALTEN (-ClientOnly):
    Löst NUR die Client-Aktionszyklen auf den angegebenen Computern aus.

.PARAMETER ComputerName
    (Obligatorisch) Ein oder mehrere Hostnamen (durch Komma getrennt), auf denen die
    Client-Zyklen ausgelöst werden sollen.

.PARAMETER ADGroupName
    (Optional) Überspringt die interaktive GUI-Auswahl und verwendet stattdessen die
    angegebene AD-Gruppe für den Server-Sync.

.PARAMETER ClientOnly
    (Optional) Ein Switch-Parameter, der das Skript anweist, NUR die Client-Zyklen
    auszulösen und alle Server-Synchronisierungsschritte zu überspringen.

.PARAMETER Silent
    (Optional) Ein Switch-Parameter, der die standardmäßig aktivierte ausführliche
    Ausgabe (Write-Verbose) unterdrückt.

.EXAMPLE
    .\Invoke-SCCMFullSync.ps1 -ComputerName 'CLIENT01'

    Öffnet das GUI-Auswahlfenster für die AD-Gruppe. Nach Auswahl wird der Full-Sync
    (Server + Client) für 'CLIENT01' ausgeführt.

.EXAMPLE
    .\Invoke-SCCMFullSync.ps1 -ComputerName 'CLIENT01','CLIENT02'

    Öffnet das GUI-Auswahlfenster. Nach Auswahl wird der Full-Sync (Server) ausgeführt
    und anschließend die Client-Zyklen auf 'CLIENT01' UND 'CLIENT02' ausgelöst.

.EXAMPLE
    .\Invoke-SCCMFullSync.ps1 -ComputerName 'CLIENT01' -ADGroupName 'APP - Adobe Install (AD)'

    Überspringt die GUI-Auswahl und führt den Full-Sync für die Gruppe
    'APP - Adobe Install (AD)' und den Client 'CLIENT01' aus.

.EXAMPLE
    .\Invoke-SCCMFullSync.ps1 -ComputerName 'CLIENT01' -ClientOnly

    Löst NUR die Client-Zyklen auf 'CLIENT01' aus.

.NOTES
    Autor: PS-Coding Assistant
    Version: 1.5 (Zusammengeführt und Multi-Host-fähig)
    Datum: 2025-10-28
    
    WICHTIGE VORAUSSETZUNGEN (Standardmodus):
    1. Die SCCM Admin Konsole (und 'ConfigurationManager' Modul) muss installiert sein.
    2. Das ausführende Konto benötigt SCCM-Admin-Rechte.
    3. Der ausführende PC muss ein SCCM-Client und Domänenmitglied sein.
#>

[CmdletBinding(DefaultParameterSetName='FullSync')]
param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage='Ein oder mehrere Zielcomputer-Hostnamen')]
    [string[]]$ComputerName,

    [Parameter(Mandatory=$true, ParameterSetName='ClientOnly', HelpMessage='Nur Client-Zyklen auslösen')]
    [switch]$ClientOnly,

    [Parameter(Mandatory=$false, ParameterSetName='FullSync', HelpMessage='AD-Gruppe direkt angeben (überspringt GUI)')]
    [string]$ADGroupName,

    [Parameter(Mandatory=$false)]
    [switch]$Silent
)

# --- 1. Konfiguration und Vorbereitung ---

# Definition der SCCM Client-Aktionen (GUIDs und Anzeigenamen)
$SCCMCycles = @(
    @{ Name = "Computerrichtlinie abrufen & evaluieren"; Guid = "{00000000-0000-0000-0000-000000000001}" },
    @{ Name = "Software-Inventur-Zyklus"; Guid = "{00000000-0000-0000-0000-000000000002}" },
    @{ Name = "Hardware-Inventur-Zyklus"; Guid = "{00000000-0000-0000-0000-000000000007}" },
    @{ Name = "Anwendungsbereitstellungs-Evaluation"; Guid = "{00000000-0000-0000-0000-000000000021}" },
    @{ Name = "Software-Updates-Scan-Zyklus"; Guid = "{00000000-0000-0000-0000-000000000108}" }
)
$WmiNamespace = 'root\ccm'
$WmiClass = 'SMS_Client'

# Globale Variablen für Server-Details
$Global:SCCMSiteCode = $null
$Global:SCCMSiteServer = $null
$Global:SCCMWmiSiteNamespace = $null

# Lokalisierung: Setzt das aktuelle Gebietsschema auf Deutsch (Deutschland).
try {
    $CultureGerman = New-Object System.Globalization.CultureInfo("de-DE")
    [System.Threading.Thread]::CurrentThread.CurrentUICulture = $CultureGerman
    [System.Threading.Thread]::CurrentThread.CurrentCulture = $CultureGerman
} catch { Write-Warning "Konnte Kultur 'de-DE' nicht setzen." }

# Verbose-Ausgabe standardmäßig aktivieren
if (-not $Silent) {
    $VerbosePreference = 'Continue'
    Write-Verbose "⚙️ Skript-Konfiguration: Verbose-Modus aktiviert (Standard). Kultur: $($CultureGerman.Name)."
} else {
    $VerbosePreference = 'SilentlyContinue'
}


# --- 2. Server-Synchronisierung (Standard: FullSync) ---

if (-not $ClientOnly) {
    Write-Host "`n--- Teil 1: SCCM Server-Synchronisierung (FullSync) ---"
    Write-Verbose "Modus: FullSync. Starte dynamische Ermittlung und Server-Trigger."

    try {
        # --- 2a. Dynamische Ermittlung des Site Servers ---
        Write-Host "⚙️ 1a. Ermittle SCCM SiteCode und SiteServer dynamisch..."
        
        Write-Verbose "Lese WMI (root\CCM\SMS_Client) des lokalen PCs..."
        $SiteCode = (Get-WmiObject -Namespace "root\CCM" -Class "SMS_Client" -Property AssignedSiteCode -ErrorAction Stop | Select-Object -ExpandProperty AssignedSiteCode)
        if (-not $SiteCode) { throw "Konnte den zugewiesenen SiteCode nicht vom lokalen SCCM-Client (WMI) lesen." }
        $Global:SCCMSiteCode = $SiteCode
        Write-Verbose "Lokaler SiteCode gefunden: $SiteCode"

        Write-Verbose "Durchsuche AD (LDAP) nach dem SiteServer FQDN für SiteCode '$SiteCode'..."
        $Searcher = New-Object System.DirectoryServices.DirectorySearcher
        $Searcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://CN=System Management,CN=System,$((Get-ADDomain).DistinguishedName)")
        $Searcher.Filter = "(&(objectClass=mSSMSSite)(cn=$SiteCode))"
        $Searcher.PropertiesToLoad.Add("mSSMSServerName") | Out-Null
        $Result = $Searcher.FindOne()
        
        if ($Result) {
            $SiteServer = $Result.Properties["mssmsservername"][0]
            $Global:SCCMSiteServer = $SiteServer
            Write-Host "✅ SiteServer dynamisch ermittelt: $SiteServer"
        } else {
            throw "Konnte keinen SiteServer FQDN für SiteCode '$SiteCode' im AD (System Management Container) finden."
        }
        $Global:SCCMWmiSiteNamespace = "root\SMS\site_$($Global:SCCMSiteCode)"

        # --- 2b. SCCM Modul-Import und Verbindung ---
        $AdminConsolePath = $env:SMS_ADMIN_UI_PATH
        if (-not $AdminConsolePath) { $AdminConsolePath = "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole" }
        $ModulePath = Join-Path -Path $AdminConsolePath -ChildPath "bin\ConfigurationManager.psd1"
        if (-not (Test-Path $ModulePath)) { throw "Das SCCM PowerShell Modul (ConfigurationManager.psd1) wurde unter '$ModulePath' nicht gefunden. Ist die SCCM Admin Konsole installiert?" }
        
        Write-Verbose "Importiere SCCM-Modul von: $ModulePath"
        Import-Module $ModulePath -ErrorAction Stop
        
        Write-Verbose "Wechsle zum PSDrive: $($Global:SCCMSiteCode):"
        Set-Location "$($Global:SCCMSiteCode):" -ErrorAction Stop

        # --- 2c. AD-Gruppenauswahl (Interaktiv oder per Parameter) ---
        $ADGroupList = @()
        
        if (-not [string]::IsNullOrEmpty($ADGroupName)) {
            Write-Verbose "Verwende AD-Gruppe aus Parameter: $ADGroupName"
            $ADGroupList = @($ADGroupName)
        } else {
            Write-Host "⚙️ 1b. Suche verfügbare AD-Gruppen mit Bereitstellungen (kann dauern)..."
            
            # Hole alle Sammlungs-IDs, die *irgendeine* Bereitstellung haben
            $DeployedCollectionIDs = (Get-CMDeployment | Select-Object -ExpandProperty CollectionID) | Sort-Object -Unique
            Write-Verbose "Analysiere $($DeployedCollectionIDs.Count) Sammlungen auf AD-Gruppen-Regeln..."
            
            $Regex = 'SecurityGroupName\s*=\s*["'']([^"'']*)["'']'
            $FoundGroups = foreach ($ID in $DeployedCollectionIDs) {
                $Collection = Get-CMCollection -Id $ID -ErrorAction SilentlyContinue
                if (-not $Collection) { continue }
                $QueryRules = Get-CMCollectionMembershipRule -CollectionId $ID -RuleType Query -ErrorAction SilentlyContinue
                
                foreach ($Rule in $QueryRules) {
                    if ($Rule.QueryExpression -match $Regex) {
                        $Matches = [regex]::Matches($Rule.QueryExpression, $Regex)
                        foreach ($Match in $Matches) {
                            $RawGroupName = $Match.Groups[1].Value
                            [PSCustomObject]@{
                                ADGroupName    = $RawGroupName.Split('\')[-1] # Nimm nur den Namen
                                CollectionName = $Collection.Name
                            }
                        }
                    }
                }
            }
            
            $UniqueGroups = $FoundGroups | Sort-Object ADGroupName -Unique
            
            if (-not $UniqueGroups) { throw "Keine AD-Gruppen gefunden, die Bereitstellungen zugeordnet sind." }

            Write-Host "Bitte wählen Sie die gewünschte(n) AD-Gruppe(n) aus dem GUI-Fenster..."
            $Selection = $UniqueGroups | Out-GridView -Title "AD-Gruppe(n) für Sync auswählen" -PassThru
            
            if ($Selection) {
                $ADGroupList = $Selection.ADGroupName
            } else {
                throw "Keine Gruppe im GUI ausgewählt. Breche Server-Synchronisierung ab."
            }
        }

        # --- 2d. AD Group Discovery auslösen (Einmal) ---
        Write-Host "⚙️ 1c. Erzwinge 'Active Directory Group Discovery' auf $SiteServer..."
        $ADGroupDiscovery = Get-WmiObject -Namespace $Global:SCCMWmiSiteNamespace -Class "SMS_DiscoveryMethods" -Filter "MethodName = 'Active Directory Group Discovery'" -ComputerName $SiteServer -ErrorAction Stop
        Invoke-WmiMethod -InputObject $ADGroupDiscovery -Name "DiscoverNow" -ErrorAction Stop
        Write-Host "✅ AD Group Discovery erfolgreich angestoßen."

        # --- 2e. Sammlungs-Aktualisierung (Looping) ---
        Write-Host "⚙️ 1d. Erzwinge Sammlungs-Updates für ausgewählte Gruppen..."
        foreach ($Group in $ADGroupList) {
            Write-Host "   -> Suche Sammlungen, die '$Group' verwenden..."
            
            $Query = "SELECT * FROM SMS_Collection WHERE CollectionRules WQL LIKE '%$Group%'"
            $WmiCollections = Get-WmiObject -Namespace $Global:SCCMWmiSiteNamespace -Query $Query -ComputerName $SiteServer -ErrorAction Stop
            
            if (-not $WmiCollections) {
                Write-Warning "   -> Keine Sammlung gefunden, die '$Group' in einer WQL-Regel verwendet."
                continue
            }
            
            foreach ($Collection in $WmiCollections) {
                Write-Host "     -> Erzwinge Update für: '$($Collection.Name)'"
                Invoke-CMCollectionUpdate -CollectionName $Collection.Name -ErrorAction Stop
            }
        }
        
        Write-Host "✅ Server-Sync: Alle relevanten Sammlungen wurden angestoßen."
        Set-Location "C:"
    }
    catch {
        Write-Error "🚨 SCHWERER FEHLER (Server-Sync): $($_.Exception.Message)"
        Write-Error "   Stellen Sie sicher, dass die SCCM Konsole installiert ist, Sie SCCM-Admin-Rechte haben,"
        Write-Error "   und der PC ein Domänenmitglied/SCCM-Client ist."
        Write-Host "   Fahre mit Client-Trigger fort, aber die Richtlinie ist möglicherweise veraltet."
    }
} else {
    Write-Verbose "Modus: ClientOnly. Überspringe Server-Synchronisierung."
}


# --- 3. Client-Trigger (für alle Hosts) ---

Write-Host "`n--- Teil 2: SCCM Client-Trigger ---"

# Schleife durch jeden Computer, der im -ComputerName Parameter angegeben wurde
foreach ($Computer in $ComputerName) {
    Write-Host "`n----- Verarbeite Client: $($Computer) -----"
    
    try {
        # --- 3a. Erreichbarkeitsprüfung (Client) ---
        Write-Host "📡 Prüfe Erreichbarkeit von Client: $($Computer)..."
        Write-Verbose "Führe Test-Connection mit Zählzahl 1 aus."
        if (-not (Test-Connection -ComputerName $Computer -Count 1 -Quiet -ErrorAction Stop)) {
            throw "Ping fehlgeschlagen oder geblockt."
        }
        Write-Host "✅ Client ist erreichbar."
        Write-Verbose "Erreichbarkeit erfolgreich geprüft. Beginne mit WMI-Aufrufen am Client..."

        # --- 3b. Sequenzielles Auslösen aller Client-Zyklen ---
        Write-Host "🚀 Starte das Auslösen der SCCM-Client-Zyklen auf $($Computer)..."
        foreach ($Cycle in $SCCMCycles) {
            $CycleName = $Cycle.Name
            $CycleGuid = $Cycle.Guid
            Write-Host "  -> ⚙️ Löst '$($CycleName)' aus..."
            Write-Verbose "WMI-Aufruf (Client): Klasse='$WmiClass', Methode='TriggerSchedule', GUID='$CycleGuid'."
            
            Invoke-WmiMethod -ComputerName $Computer -Namespace $WmiNamespace -Class $WmiClass -Name 'TriggerSchedule' -ArgumentList $CycleGuid -ErrorAction Stop
            Write-Host "     ✅ $CycleName erfolgreich an Client übergeben."
        }
        Write-Host "ℹ️ Client-Trigger für **$($Computer)** abgeschlossen."
    }
    catch {
        Write-Error "🚨 FEHLER bei Client '$($Computer)': $($_.Exception.Message)"
        Write-Host "   -> Überspringe diesen Host und fahre mit dem nächsten fort (falls vorhanden)."
        continue
    }
}

# --- 4. Abschluss und Hinweis ---

$CurrentTimeGerman = Get-Date -Format 'dd.MM.yyyy HH:mm:ss'
Write-Host "`n--- Skript-Ende [$($CurrentTimeGerman)] ---"
Write-Host "Alle Aktionen wurden für die angegebenen Hosts und Gruppen ausgeführt."
