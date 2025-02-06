<#
.SYNOPSIS
    Erstellt AD-Benutzer als exakte Kopie eines Vorlagenbenutzers
.DESCRIPTION
    Dieses Skript erstellt neue AD-Benutzer durch Klonen eines vorhandenen Benutzers.
    Es beinhetzt Passwort-Komplexitätsprüfungen und deutsche Lokalisierung.
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(ParameterSetName='CSV')]
    [ValidateScript({Test-Path $_})]
    [string]$CsvPfad,

    [Parameter(ParameterSetName='Single')]
    [string]$VorlagenBenutzer,

    [Parameter(ParameterSetName='Single')]
    [string]$NeuerBenutzer,

    [Parameter(ParameterSetName='Single')]
    [string]$Kennwort,

    [switch]$Verifizieren
)

begin {
    #region Hilfsfunktionen
    function Test-PasswortKomplexität {
        param([string]$Passwort)
        $regeln = @(
            { $Passwort.Length -ge 12 },
            { $Passwort -cmatch '[A-Z]' },    # Mind. 1 Großbuchstabe
            { $Passwort -cmatch '[a-z]' },    # Mind. 1 Kleinbuchstabe
            { $Passwort -match '\d' },        # Mind. 1 Zahl
            { $Passwort -match '[\W_]' }      # Mind. 1 Sonderzeichen
        )
        return ($regeln | Where-Object { -not (& $_) }).Count -eq 0
    }

    function Schreibe-Protokoll {
        param(
            [string]$Meldung,
            [ValidateSet('INFO','WARNUNG','FEHLER')]
            [string]$Stufe = 'INFO'
        )
        $zeitstempel = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
        Write-Verbose "[$zeitstempel][$Stufe] $Meldung"
    }
    #endregion

    # Initialisierung
    Import-Module ActiveDirectory -ErrorAction Stop
    $global:Protokolldatei = Join-Path $PSScriptRoot "ADErstellung_$(Get-Date -Format 'yyyyMMdd').log"
}

process {
    try {
        # Passwortkomplexität prüfen
        if ($PSBoundParameters.ContainsKey('Kennwort') -and -not (Test-PasswortKomplexität $Kennwort)) {
            throw "Passwort entspricht nicht den Sicherheitsrichtlinien!"
        }

        # CSV-Verarbeitung (deutsches Format)
        if ($PSCmdlet.ParameterSetName -eq 'CSV') {
            $daten = Import-Csv -Path $CsvPfad -Delimiter ';' -Encoding UTF8
            foreach ($eintrag in $daten) {
                if (-not (Test-PasswortKomplexität $eintrag.Kennwort)) {
                    Schreibe-Protokoll "Ungültiges Passwort für $($eintrag.NeuerBenutzer)" -Stufe WARNUNG
                    continue
                }

                # Benutzererstellung
                $securePass = ConvertTo-SecureString $eintrag.Kennwort -AsPlainText -Force
                $params = @{
                    SamAccountName = $eintrag.NeuerBenutzer
                    AccountPassword = $securePass
                    Instance = (Get-ADUser $eintrag.VorlagenBenutzer -Properties *)
                }

                if ($PSCmdlet.ShouldProcess($eintrag.NeuerBenutzer, "Erstelle Benutzer")) {
                    New-ADUser @params -PassThru | Out-Null
                    Schreibe-Protokoll "Benutzer $($eintrag.NeuerBenutzer) erstellt"
                }
            }
        }
        else {
            # Interaktiver Modus
            if ([string]::IsNullOrEmpty($VorlagenBenutzer)) {
                $VorlagenBenutzer = Read-Host "Vorlagenbenutzer"
            }

            # ... (restliche Logik)
        }
    }
    catch {
        Schreibe-Protokoll "Fehler: $_" -Stufe FEHLER
        exit 1
    }
}
