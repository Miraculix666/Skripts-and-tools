<#
.SYNOPSIS
    Erstellt AD-Benutzer als Kopie eines Vorlagenbenutzers mit deutscher Lokalisierung
.DESCRIPTION
    Dieses Skript erstellt neue AD-Benutzer als Klone eines vorhandenen Benutzers.
    Es unterstützt drei Modi:
    1. Interaktiv (keine Parameter)
    2. Einzelbenutzer-Erstellung via Parameter
    3. Batch-Erstellung via CSV-Datei (Deutsches Format mit Semikolon-Trenner)
    
    Version: 5.1-compatible
    Autor: IT-Support
    Letzte Änderung: 06.02.2025
#>

[CmdletBinding(DefaultParameterSetName='Interactive', SupportsShouldProcess=$true)]
param(
    [Parameter(ParameterSetName='CSV', Mandatory=$true)]
    [ValidateScript({Test-Path $_ -PathTypeLeaf})]
    [string]$CsvPfad,
    
    [Parameter(ParameterSetName='Single')]
    [string]$VorlagenBenutzer,
    
    [Parameter(ParameterSetName='Single')]
    [string]$NeuerBenutzer,
    
    [Parameter(ParameterSetName='Single')]
    [string]$Kennwort,
    
    [Parameter()]
    [switch]$Verifizieren,
    
    [Parameter()]
    [switch]$Verbose
)

begin {
    #region Hilfsfunktionen
    function Schreibe-Protokoll {
        param(
            [string]$Meldung,
            [ValidateSet('INFO','WARNUNG','FEHLER','ERFOLG')]
            [string]$Stufe = 'INFO'
        )
        $zeitstempel = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
        $eintrag = "[$zeitstempel][$Stufe] $Meldung"
        Add-Content -Path $Protokolldatei -Value $eintrag
        
        if ($Verbose) {
            $farbe = switch ($Stufe) {
                'INFO'    { 'Gray' }
                'WARNUNG' { 'Yellow' }
                'FEHLER'  { 'Red' }
                'ERFOLG'  { 'Green' }
            }
            Write-Host $eintrag -ForegroundColor $farbe
        }
    }

    function Prüfe-KennwortKomplexität {
        param([string]$Kennwort)
        $regeln = @(
            { $Kennwort.Length -ge 12 },
            { $Kennwort -cmatch '[A-Z]' },
            { $Kennwort -cmatch '[a-z]' },
            { $Kennwort -match '\d' },
            { $Kennwort -match '[\W_]' }
        )
        return ($regeln | Where-Object { -not (& $_) }).Count -eq 0
    }
    #endregion

    # Initialisierung
    $Protokolldatei = Join-Path $PSScriptRoot "ADBenutzerErstellung_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    Import-Module ActiveDirectory -ErrorAction Stop
}

process {
    try {
        #region Parameterverarbeitung
        switch ($PSCmdlet.ParameterSetName) {
            'Interactive' {
                $VorlagenBenutzer = Read-Host "Vorlagenbenutzer (SAMAccountName)"
                $NeuerBenutzer = Read-Host "Neuer Benutzername"
                $Kennwort = Read-Host "Kennwort" -AsSecureString
                $Kennwort = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
                    [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Kennwort)
                )
            }
            'CSV' {
                $benutzerListe = Import-Csv -Path $CsvPfad -Delimiter ';' -Encoding UTF8
                # CSV-Validierung
                if (-not ($benutzerListe[0].PSObject.Properties.Name -contains 'VorlagenBenutzer' -and
                        $benutzerListe[0].PSObject.Properties.Name -contains 'NeuerBenutzer' -and
                        $benutzerListe[0].PSObject.Properties.Name -contains 'Kennwort')) {
                    throw "CSV muss folgende Spalten enthalten: VorlagenBenutzer;NeuerBenutzer;Kennwort"
                }
            }
        }
        #endregion

        #region Benutzererstellung
        if ($PSCmdlet.ParameterSetName -eq 'CSV') {
            foreach ($eintrag in $benutzerListe) {
                Schreibe-Protokoll "Verarbeite Benutzer: $($eintrag.NeuerBenutzer)" -Stufe INFO
                
                if (-not (Prüfe-KennwortKomplexität -Kennwort $eintrag.Kennwort)) {
                    Schreibe-Protokoll "Kennwort entspricht nicht den Richtlinien" -Stufe WARNUNG
                    continue
                }

                # Konvertierung zu SecureString
                $secureKennwort = ConvertTo-SecureString $eintrag.Kennwort -AsPlainText -Force

                # Benutzerklonung
                $vorlage = Get-ADUser -Identity $eintrag.VorlagenBenutzer -Properties *
                $neuerBenutzerParams = @{
                    Instance           = $vorlage
                    SamAccountName     = $eintrag.NeuerBenutzer
                    UserPrincipalName  = "$($eintrag.NeuerBenutzer)@$((Get-ADDomain).DNSRoot)"
                    AccountPassword    = $secureKennwort
                    Enabled            = $true
                    Path               = ($vorlage.DistinguishedName -split ',',2)[1]
                }

                if ($PSCmdlet.ShouldProcess($eintrag.NeuerBenutzer, "Benutzer erstellen")) {
                    $neuerBenutzer = New-ADUser @neuerBenutzerParams -PassThru
                    Schreibe-Protokoll "Benutzer $($neuerBenutzer.SamAccountName) erfolgreich erstellt" -Stufe ERFOLG
                }
            }
        }
        else {
            # Einzelbenutzer-Erstellung
            if (-not (Prüfe-KennwortKomplexität -Kennwort $Kennwort)) {
                throw "Kennwort entspricht nicht den Komplexitätsregeln"
            }

            $secureKennwort = ConvertTo-SecureString $Kennwort -AsPlainText -Force
            $vorlage = Get-ADUser -Identity $VorlagenBenutzer -Properties *

            $neuerBenutzerParams = @{
                Instance           = $vorlage
                SamAccountName     = $NeuerBenutzer
                UserPrincipalName  = "$NeuerBenutzer@$((Get-ADDomain).DNSRoot)"
                AccountPassword    = $secureKennwort
                Enabled            = $true
                Path               = ($vorlage.DistinguishedName -split ',',2)[1]
            }

            if ($PSCmdlet.ShouldProcess($NeuerBenutzer, "Benutzer erstellen")) {
                $neuerBenutzer = New-ADUser @neuerBenutzerParams -PassThru
                Schreibe-Protokoll "Benutzer $NeuerBenutzer erfolgreich erstellt" -Stufe ERFOLG
            }
        }
        #endregion

        #region Verifikation
        if ($Verifizieren) {
            if ($neuerBenutzer) {
                $vergleich = Compare-Object $vorlage $neuerBenutzer -Property SamAccountName,Enabled,PasswordLastSet
                if ($vergleich) {
                    Schreibe-Protokoll "Abweichungen gefunden:`n$($vergleich | Out-String)" -Stufe WARNUNG
                }
            }
        }
        #endregion
    }
    catch {
        Schreibe-Protokoll "KRITISCHER FEHLER: $_" -Stufe FEHLER
        exit 1
    }
}

end {
    Schreibe-Protokoll "Prozess abgeschlossen. Details in: $Protokolldatei" -Stufe INFO
}
