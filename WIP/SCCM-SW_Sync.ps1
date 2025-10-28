<#
.SYNOPSIS
    Führt standardmäßig eine vollständige SCCM-Synchronisierung (AD-Erkennung, Sammlungs-
    aktualisierung und Client-Zyklen) aus, um Software-Bereitstellungen zu beschleunigen.

.DESCRIPTION
    Dieses Skript ist ein All-in-One-Tool zur Beschleunigung von SCCM-Bereitstellungen,
    insbesondere in Umgebungen mit Identitätsmanagern (wie Dell One Identity), die
    Bereitstellungen über AD-Gruppen steuern.

    STANDARDVERHALTEN (FullSync):
    1. Löst die "Active Directory Group Discovery" auf dem Site Server aus.
    2. Löst die "Collection Membership Evaluation" für eine bestimmte Sammlung aus.
    3. Löst alle wichtigen Client-Aktionszyklen auf dem Zielcomputer aus.

    ALTERNATIVES VERHALTEN (-ClientOnly):
    Löst NUR die Client-Aktionszyklen auf dem Zielcomputer aus.

.PARAMETER SiteServer
    (Standard) Der FQDN des SCCM Primary Site Servers (z.B. sccm-ps01.domain.local).
    Im Standardmodus (FullSync) erforderlich.

.PARAMETER CollectionName
    (Standard) Der genaue Name der SCCM-Sammlung (z.B. "APP - MyApp 1.0 Install"),
    deren Mitgliedschaft aktualisiert werden muss. Im Standardmodus (FullSync) erforderlich.

.PARAMETER ComputerName
    Der Hostname oder die IP-Adresse des Zielcomputers (Client).
    Wenn nicht angegeben, wird interaktiv nach dem Namen gefragt.

.PARAMETER ClientOnly
    (Optional) Ein Switch-Parameter, der das Skript anweist, NUR die Client-Zyklen
    auszulösen und alle Server-Synchronisierungsschritte zu überspringen.

.PARAMETER Silent
    (Optional) Ein Switch-Parameter, der die standardmäßig aktivierte ausführliche
    Ausgabe (Write-Verbose) unterdrückt.

.EXAMPLE
    .\Invoke-SCCMFullSync.ps1 -SiteServer 'sccm.domain.local' -CollectionName 'APP - Adobe Install' -ComputerName 'CLIENT01'

    Führt die vollständige Synchronisierung aus: AD Discovery -> Collection Update -> Client Trigger auf CLIENT01.

.EXAMPLE
    .\Invoke-SCCMFullSync.ps1 -SiteServer 'sccm.domain.local' -CollectionName 'APP - Adobe Install'

    Führt die vollständige Server-Synchronisierung aus und fragt interaktiv nach dem Client-ComputerName.

.EXAMPLE
    .\Invoke-SCCMFullSync.ps1 -ClientOnly -ComputerName 'CLIENT01'

    Löst NUR die Client-Zyklen auf CLIENT01 aus.

.NOTES
    Autor: PS-Coding Assistant
    Version: 1.3 (FullSync als Standard; AD Group Discovery hinzugefügt)
    Datum: 2025-10-28
    
    WICHTIGE VORAUSSETZUNGEN FÜR SERVER-SYNC (Standardmodus):
    1. Die SCCM Admin Konsole (und das 'ConfigurationManager' Modul) muss auf dem
       Computer installiert sein, der dieses Skript ausführt.
    2. Das ausführende Benutzerkonto benötigt SCCM-Administratorrechte.
#>

[CmdletBinding(DefaultParameterSetName='FullSync')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='FullSync', HelpMessage='FQDN des SCCM Primary Site Servers')]
    [string]$SiteServer,

    [Parameter(Mandatory=$true, ParameterSetName='FullSync', HelpMessage='Name der zu aktualisierenden Sammlung')]
    [string]$CollectionName,

    [Parameter(Mandatory=$true, ParameterSetName='ClientOnly', HelpMessage='Nur Client-Zyklen auslösen')]
    [switch]$ClientOnly,
    
    [Parameter(Mandatory=$false, Position=0, HelpMessage='Hostname oder IP des Zielcomputers (Client)')]
    [string]$ComputerName,

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

# Lokalisierung: Setzt das aktuelle Gebietsschema auf Deutsch (Deutschland).
$CultureGerman = New-Object System.Globalization.CultureInfo("de-DE")
[System.Threading.Thread]::CurrentThread.CurrentUICulture = $CultureGerman
[System.Threading.Thread]::CurrentThread.CurrentCulture = $CultureGerman

# Verbose-Ausgabe standardmäßig aktivieren
if (-not $Silent) {
    $VerbosePreference = 'Continue'
    Write-Verbose "⚙️ Skript-Konfiguration: Verbose-Modus aktiviert (Standard). Kultur: $($CultureGerman.Name)."
} else {
    $VerbosePreference = 'SilentlyContinue'
}


# --- 2. Server-Synchronisierung (Standard: FullSync) ---

if ($PSCmdlet.ParameterSetName -eq 'FullSync') {
    Write-Host "`n--- Teil 1: SCCM Server-Synchronisierung (FullSync) ---"
    Write-Verbose "Modus: FullSync. Führe AD Discovery und Collection Update aus."

    try {
        # --- 2a. SCCM Modul-Import und Verbindung ---
        $AdminConsolePath = $env:SMS_ADMIN_UI_PATH
        if (-not $AdminConsolePath) {
            $AdminConsolePath = "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole"
        }
        $ModulePath = Join-Path -Path $AdminConsolePath -ChildPath "bin\ConfigurationManager.psd1"

        if (-not (Test-Path $ModulePath)) {
            throw "Das SCCM PowerShell Modul (ConfigurationManager.psd1) wurde unter '$ModulePath' nicht gefunden. Ist die SCCM Admin Konsole installiert?"
        }

        Write-Verbose "Importiere SCCM-Modul von: $ModulePath"
        Import-Module $ModulePath -ErrorAction Stop

        $SiteCode = (Get-WmiObject -Namespace "root\SMS" -Class "SMS_ProviderLocation" -ComputerName $SiteServer -ErrorAction Stop | Select-Object -First 1).SiteCode
        if (-not $SiteCode) {
            throw "Konnte den SiteCode für $SiteServer nicht ermitteln."
        }
        
        Write-Verbose "Verbinde mit SiteCode $SiteCode (ermittelt von $SiteServer)"
        Set-Location "$($SiteCode):" -ErrorAction Stop
        $WmiSiteNamespace = "root\SMS\site_$($SiteCode)"

        # --- 2b. AD Group Discovery auslösen ---
        Write-Host "⚙️ 1a. Erzwinge 'Active Directory Group Discovery' auf $SiteServer..."
        Write-Verbose "Suche WMI-Discovery-Methode 'Active Directory Group Discovery' im Namespace '$WmiSiteNamespace'"
        
        $ADGroupDiscovery = Get-WmiObject -Namespace $WmiSiteNamespace -Class "SMS_DiscoveryMethods" -Filter "MethodName = 'Active Directory Group Discovery'" -ComputerName $SiteServer -ErrorAction Stop
        
        if ($ADGroupDiscovery) {
            Invoke-WmiMethod -InputObject $ADGroupDiscovery -Name "DiscoverNow" -ErrorAction Stop
            Write-Host "✅ AD Group Discovery erfolgreich angestoßen."
        } else {
            Write-Warning "Konnte die 'Active Directory Group Discovery'-Methode nicht finden."
        }

        # --- 2c. Sammlungs-Aktualisierung auslösen ---
        Write-Verbose "Suche nach Sammlung '$CollectionName'..."
        $Collection = Get-CMCollection -Name $CollectionName -ErrorAction Stop
        
        if (-not $Collection) {
            throw "Sammlung '$CollectionName' konnte nicht gefunden werden."
        }

        Write-Host "⚙️ 1b. Erzwinge Aktualisierung der Sammlungsmitgliedschaft für: '$($Collection.Name)'"
        Invoke-CMCollectionUpdate -CollectionName $Collection.Name -ErrorAction Stop
        
        Write-Host "✅ Server-Sync: Aktualisierung für '$($Collection.Name)' erfolgreich angestoßen."
        
        Set-Location "C:"
    }
    catch {
        Write-Error "🚨 SCHWERER FEHLER (Server-Sync): $($_.Exception.Message)"
        Write-Error "   Stellen Sie sicher, dass die SCCM Konsole installiert ist und Sie ausreichende Rechte haben."
        Write-Host "   Fahre mit Client-Trigger fort, aber die Richtlinie ist möglicherweise veraltet."
    }
} else {
    Write-Verbose "Modus: ClientOnly. Überspringe Server-Synchronisierung."
}


# --- 3. Parameter-Validierung (Client) ---

if ([string]::IsNullOrEmpty($ComputerName)) {
    Write-Verbose "Der Parameter 'ComputerName' fehlt. Fordere interaktive Eingabe an."
    do {
        $ComputerName = Read-Host -Prompt "Bitte geben Sie den Hostnamen oder die IP des ZIELCOMPUTERS ein"
    } until (-not [string]::IsNullOrEmpty($ComputerName))
}

Write-Verbose "Zielcomputer (Client): **$($ComputerName)**."
Write-Host "`n--- Teil 2: SCCM Client-Trigger ---"


# --- 4. Erreichbarkeitsprüfung (Client) ---

Write-Host "📡 Prüfe Erreichbarkeit von Client: $($ComputerName)..."
Write-Verbose "Führe Test-Connection mit Zählzahl 1 aus."

try {
    if (-not (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet -ErrorAction Stop)) {
        throw "Ping fehlgeschlagen oder geblockt."
    }
    Write-Host "✅ Client ist erreichbar."
    Write-Verbose "Erreichbarkeit erfolgreich geprüft. Beginne mit WMI-Aufrufen am Client..."
}
catch {
    Write-Error "🚨 Fehler: Client-Hostname '$($ComputerName)' ist nicht erreichbar oder WMI-Zugriff nicht möglich. Details: $($_.Exception.Message)"
    exit 1
}


# --- 5. Sequenzielles Auslösen aller Client-Zyklen ---

Write-Host "`n🚀 Starte das Auslösen der SCCM-Client-Zyklen auf $($ComputerName)..."

foreach ($Cycle in $SCCMCycles) {
    $CycleName = $Cycle.Name
    $CycleGuid = $Cycle.Guid
    
    Write-Host "  -> ⚙️ Löst '$($CycleName)' aus..."
    Write-Verbose "WMI-Aufruf (Client): Klasse='$WmiClass', Methode='TriggerSchedule', GUID='$CycleGuid'."

    try {
        Invoke-WmiMethod -ComputerName $ComputerName -Namespace $WmiNamespace -Class $WmiClass -Name 'TriggerSchedule' -ArgumentList $CycleGuid -ErrorAction Stop
        Write-Host "     ✅ $CycleName erfolgreich an Client übergeben."
    }
    catch {
        Write-Error "     🚨 Fehler beim Auslösen von '$($CycleName)' (Client): $($_.Exception.Message)."
    }
}


# --- 6. Abschluss und Hinweis ---

$CurrentTimeGerman = Get-Date -Format 'dd.MM.yyyy HH:mm:ss'
Write-Host "`n--- Skript-Ende [$($CurrentTimeGerman)] ---"
Write-Host "ℹ️ Alle relevanten Zyklen wurden an **$($ComputerName)** gesendet."
Write-Host "   Die neue Software/Aktualisierung sollte nun im Software Center erscheinen und ausgeführt werden."
