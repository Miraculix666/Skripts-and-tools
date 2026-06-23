# FileName: PS_LAccount_Apply.ps1
# Version:  1.0
# Beschreibung: Liest die Properties-CSV des AD-Sync-Managers (v9.x) und
#               wendet Aenderungen (Property-Updates, OU-Verschiebungen,
#               Gruppen-Zuweisungen, Loeschungen) auf Active Directory an.
#
# EINGABE:  L-Kennungen_Properties_vX.X.csv  (aus PS_LAccount_Manager v9.x)
# AUSGABE:  Apply-Log CSV  +  Log-Datei
#
# AKTIONEN (gesteuert durch Tabellenspalten):
#   NICHT_KONFORM enthält Codes  -> Set-ADUser fuer die betroffenen Attribute
#   Codes:  VN   = GivenName (Vorname)
#           NN   = Surname   (Nachname)
#           DN   = DisplayName
#           DESC = Description
#           ORT  = l (Ort/Stadt)
#           GEB  = physicalDeliveryOfficeName (Buero)
#           DEZ  = department
#           INFO = info
#   LOESCHEN <> ""             -> Konto deaktivieren (kein Hard-Delete)
#   AENDERN_OU gefuellt        -> OU-Verschiebung pruefen / vorbereiten
#   GRP_xxx Spalten / Gruppen  -> Gruppen-Sync: X=hinzufuegen, ""=entfernen
#
# PARAMETER:
#   -CsvPath      Pfad zur Properties-CSV (Pflicht)
#   -SenderEmail  Absender fuer .eml Benachrichtigungen (Pflicht)
#   -GroupSync    Gruppen-Sync auch ausfuehren (Default: nein)
#   -Filter       Nur Zeilen mit diesen NICHT_KONFORM-Codes verarbeiten
#                 z.B. -Filter "VN,NN,DN"  -> nur Vorname/Nachname/DisplayName
#   -WhatIf       Trockenlauf: zeigt alles, aendert nichts
#   -DebugMode    Ausfuehrliches Logging
#
# BEISPIELE:
#   # Alle nicht-konformen Properties fixen, Gruppen synchronisieren:
#   .\PS_LAccount_Apply.ps1 -CsvPath .\Properties.csv -SenderEmail "it@polizei.nrw.de" -GroupSync
#
#   # Nur Vor-/Nachname und DisplayName korrigieren (Trockenlauf):
#   .\PS_LAccount_Apply.ps1 -CsvPath .\Properties.csv -SenderEmail "it@polizei.nrw.de" -Filter "VN,NN,DN" -WhatIf
#
#   # Nur Loeschungen (Deaktivierungen) ausfuehren:
#   .\PS_LAccount_Apply.ps1 -CsvPath .\Properties.csv -SenderEmail "it@polizei.nrw.de" -Filter "LOESCHEN"

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(Mandatory=$true)]
    [string] $CsvPath,

    [Parameter(Mandatory=$true)]
    [string] $SenderEmail,

    [Parameter(Mandatory=$false)]
    [switch] $GroupSync,

    [Parameter(Mandatory=$false)]
    [string] $Filter = "",          # leer = alle NICHT_KONFORM-Codes

    [Parameter(Mandatory=$false)]
    [switch] $DebugMode
)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

$Version   = "1.0"
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile   = Join-Path $ScriptDir "Apply_Log_$Timestamp.log"
$EmailDir  = Join-Path $ScriptDir "EMAILS_Apply"
$SW        = [System.Diagnostics.Stopwatch]::StartNew()

if (-not (Test-Path $EmailDir)) { New-Item -ItemType Directory -Path $EmailDir -Force | Out-Null }

# ══════════════════════════════════════════════════════════════════
#  UI & LOGGING
# ══════════════════════════════════════════════════════════════════
function Show-Banner {
    Clear-Host
    $dryRun = if ($WhatIfPreference) { "  [!! TROCKENLAUF - KEINE AENDERUNGEN !!]" } else { "" }
    Write-Host ""
    Write-Host "  +----------------------------------------------------------+" -ForegroundColor DarkCyan
    Write-Host "  |  AD L-KENNUNG APPLY MANAGER  v$Version                        |" -ForegroundColor Cyan
    Write-Host "  |  $(Get-Date -Format 'dd.MM.yyyy  HH:mm:ss')                               |" -ForegroundColor DarkCyan
    Write-Host "  |  Liest Properties-CSV und wendet Aenderungen an          |" -ForegroundColor DarkCyan
    if ($dryRun) {
    Write-Host "  |  !! TROCKENLAUF - KEINE AD-AENDERUNGEN !!                |" -ForegroundColor Yellow
    }
    Write-Host "  +----------------------------------------------------------+" -ForegroundColor DarkCyan
    Write-Host ""
}

function Write-Log {
    param(
        [string] $Msg,
        [ValidateSet("INFO","OK","WARN","ERR","DBG","STEP","SKIP","DRYRUN")] $L = "INFO",
        [switch] $NoFile
    )
    if ($L -eq "DBG" -and -not $DebugMode) { return }
    $ts      = Get-Date -Format "HH:mm:ss.fff"
    $elapsed = "+{0,8:F2}s" -f $SW.Elapsed.TotalSeconds
    $tag     = switch ($L) {
        "OK"     { "[OK]    " } "WARN"   { "[WARN]  " } "ERR"    { "[ERR]   " }
        "DBG"    { "[DBG]   " } "STEP"   { "[----]  " } "SKIP"   { "[SKIP]  " }
        "DRYRUN" { "[DRY]   " }
        default  { "[INFO]  " }
    }
    $color = switch ($L) {
        "OK"     { "Green"       } "WARN"   { "Yellow"  } "ERR"    { "Red"         }
        "DBG"    { "Magenta"     } "STEP"   { "Cyan"    } "SKIP"   { "DarkGray"    }
        "DRYRUN" { "DarkYellow"  }
        default  { "Gray"        }
    }
    $line = "[$ts]  $elapsed  $tag  $Msg"
    Write-Host $line -ForegroundColor $color
    if (-not $NoFile) { $line | Out-File $LogFile -Append }
}

function Write-Section {
    param([string]$Title)
    $bar = "-" * 62
    Write-Host ""
    Write-Host "  +$bar+" -ForegroundColor DarkGray
    Write-Host "  |  $($Title.PadRight(60))  |" -ForegroundColor White
    Write-Host "  +$bar+" -ForegroundColor DarkGray
    Write-Log "=== $Title ===" "STEP"
}

function Show-ActionPlan {
    param([array]$Plan)
    Write-Host ""
    Write-Host "  AKTIONSPLAN:" -ForegroundColor White
    Write-Host ("  {0,-16} {1,-12} {2,-35} {3}" -f "L-Kennung","Aktion","Codes / Details","Loeschen") -ForegroundColor DarkCyan
    Write-Host ("  " + ("-" * 80)) -ForegroundColor DarkGray
    foreach ($item in $Plan) {
        $color = switch ($item.Aktion) {
            "UPDATE"    { "Cyan"    }
            "DEAKTIV"   { "Red"     }
            "SKIP"      { "DarkGray"}
            "GRP-SYNC"  { "Yellow"  }
            default     { "Gray"    }
        }
        Write-Host ("  {0,-16} {1,-12} {2,-35} {3}" -f `
            $item.LID, $item.Aktion, $item.Codes, $item.Loeschen) -ForegroundColor $color
    }
    Write-Host ""
}

# ══════════════════════════════════════════════════════════════════
#  HILFSFUNKTIONEN
# ══════════════════════════════════════════════════════════════════
function Get-Str { param($v)
    if ($null -eq $v) { return "" }
    return $v.ToString().Trim()
}

# Mapping: NICHT_KONFORM-Code -> AD-Attribut-Name und Quellspalte in CSV
$CodeMap = @{
    "VN"   = @{ AdAttr = "GivenName";                   CsvCol = "AENDERN_Vorname"      }
    "NN"   = @{ AdAttr = "Surname";                     CsvCol = "AENDERN_Nachname"     }
    "DN"   = @{ AdAttr = "DisplayName";                 CsvCol = "AENDERN_DisplayName"  }
    "DESC" = @{ AdAttr = "Description";                 CsvCol = "AENDERN_Description"  }
    "ORT"  = @{ AdAttr = "City";                        CsvCol = "AENDERN_Ort"          }
    "GEB"  = @{ AdAttr = "Office";                      CsvCol = "AENDERN_Buero"        }
    "DEZ"  = @{ AdAttr = "Department";                  CsvCol = "AENDERN_Dez"          }
    "INFO" = @{ AdAttr = "OtherAttributes";             CsvCol = "AENDERN_Info"         }
}

function Send-ChangeNotification {
    param([string]$LID, [string]$Email, [string]$Changes, [string]$ActionType)
    if ([string]::IsNullOrEmpty($Email)) { return }
    $subject = "AD-Konto Aktualisierung: $LID"
    $body    = "Konto:     $LID`nAenderung: $Changes`nDatum:     $(Get-Date -Format 'dd.MM.yyyy HH:mm')`n`nDiese Nachricht wurde automatisch generiert."
    $encoded = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($subject))
    $eml = @"
From: $SenderEmail
To: $Email
Subject: =?utf-8?B?$encoded?=
MIME-Version: 1.0
Content-Type: text/plain; charset=utf-8
Content-Transfer-Encoding: 8bit

$body
"@
    $path = Join-Path $EmailDir "$($LID)_$($ActionType)_$(Get-Date -Format 'HHmmss').eml"
    $eml | Out-File -FilePath $path -Encoding utf8NoBOM
    Write-Log "EML erstellt: $(Split-Path $path -Leaf)" "DBG"
}

# ══════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════
Show-Banner
Import-Module ActiveDirectory -Verbose:$false

Write-Log "PS v$($PSVersionTable.PSVersion)  PID=$PID  WhatIf=$($WhatIfPreference)" "DBG"

# AD-Domain fuer spaetere Referenz
try {
    $ADDomain   = Get-ADDomain
    $UPNSuffix  = $ADDomain.DNSRoot
    Write-Log "AD-Domain: $UPNSuffix" "OK"
} catch {
    Write-Log "AD-Verbindung fehlgeschlagen: $_" "ERR"; exit 1
}

# ── CSV LADEN ────────────────────────────────────────────────────
Write-Section "CSV LADEN"
$CsvPath = (Resolve-Path $CsvPath).Path
Write-Log "Lese: $CsvPath" "INFO"
$t0 = $SW.Elapsed.TotalSeconds

try {
    $AllRows = Import-Csv -Path $CsvPath -Delimiter ';' -Encoding UTF8
} catch {
    # Fallback: Default-Encoding (bei ANSI-Dateien vom Sync Manager)
    $AllRows = Import-Csv -Path $CsvPath -Delimiter ';' -Encoding Default
}

Write-Log ("CSV geladen: {0} Zeilen  [{1:F2}s]" -f $AllRows.Count, ($SW.Elapsed.TotalSeconds-$t0)) "OK"

# Spalten analysieren
$AllCols   = $AllRows[0].psobject.Properties.Name
$GrpCols   = @($AllCols | Where-Object { $_ -like "GRP_*" })
$HasGruppen = ($AllCols -contains "Gruppen")   # GroupMode=Single
Write-Log ("Spalten: {0} gesamt  {1} GRP-Spalten  GroupMode-Single: {2}" -f `
    $AllCols.Count, $GrpCols.Count, $HasGruppen) "DBG"

# Filter-Codes vorbereiten
$FilterCodes = @()
if ($Filter -ne "") {
    $FilterCodes = $Filter.ToUpper() -split '[,;\s]' | Where-Object { $_ -ne "" }
    Write-Log "Filter aktiv: $($FilterCodes -join ', ')" "WARN"
}

# ── AKTIONSPLAN AUFBAUEN ─────────────────────────────────────────
Write-Section "AKTIONSPLAN analysieren"
$ActionPlan  = New-Object 'System.Collections.Generic.List[PSObject]'
$ResultLog   = New-Object 'System.Collections.Generic.List[PSObject]'

$cntUpdate  = 0
$cntDeaktiv = 0
$cntSkip    = 0
$cntGrp     = 0

foreach ($row in $AllRows) {
    $lid       = Get-Str $row."L-Kennung"
    $nkCodes   = Get-Str $row."NICHT_KONFORM"
    $loeschen  = Get-Str $row."LOESCHEN"
    $geloescht = Get-Str $row."GELOESCHT"

    # Zeile ueberspringen wenn L-Kennung fehlt
    if ([string]::IsNullOrEmpty($lid)) { continue }

    # Gruppen-Sync-Infos extrahieren
    $grpAdd    = @()
    $grpRemove = @()
    if ($GroupSync) {
        if ($GrpCols.Count -gt 0) {
            foreach ($gc in $GrpCols) {
                $grpName = $gc -replace '^GRP_',''
                $val     = Get-Str $row.$gc
                if ($val -eq "X") { $grpAdd    += $grpName }
                else              { $grpRemove += $grpName }
            }
        }
    }

    # Aktionstyp bestimmen
    $aktion    = "SKIP"
    $aktCodes  = @()
    $loeschenFlag = $false

    # Deaktivierung pruefen
    if ($loeschen -ne "") {
        $loeschenFlag = $true
        $aktion = "DEAKTIV"
        $cntDeaktiv++
    }

    # Property-Updates pruefen (nur wenn nicht schon als GELOESCHT markiert)
    if ($nkCodes -ne "" -and $geloescht -ne "XXX") {
        $rawCodes = $nkCodes -split '[,;\s]' | Where-Object { $_ -ne "" }

        # Filter anwenden
        if ($FilterCodes.Count -gt 0) {
            $aktCodes = @($rawCodes | Where-Object { $FilterCodes -contains $_ })
        } else {
            $aktCodes = $rawCodes
        }

        if ($aktCodes.Count -gt 0) {
            $aktion = "UPDATE"
            $cntUpdate++
        }
    }

    # Gruppen-Sync
    if ($GroupSync -and ($grpAdd.Count -gt 0 -or $grpRemove.Count -gt 0)) {
        if ($aktion -eq "SKIP") { $aktion = "GRP-SYNC" }
        $cntGrp++
    }

    if ($aktion -eq "SKIP") { $cntSkip++ }

    $entry = [PSCustomObject]@{
        LID           = $lid
        Aktion        = $aktion
        Codes         = ($aktCodes -join ",")
        Loeschen      = $loeschen
        GrpAdd        = $grpAdd
        GrpRemove     = $grpRemove
        Row           = $row
        LoeschenFlag  = $loeschenFlag
    }
    [void]$ActionPlan.Add($entry)
}

Write-Log ("Plan: {0} UPDATE  {1} DEAKTIV  {2} GRP-SYNC  {3} SKIP" -f `
    $cntUpdate, $cntDeaktiv, $cntGrp, $cntSkip) "INFO"

# ── AKTIONSPLAN ANZEIGEN ─────────────────────────────────────────
$toShow = @($ActionPlan | Where-Object { $_.Aktion -ne "SKIP" })
if ($toShow.Count -eq 0) {
    Write-Log "Keine Aktionen notwendig. Alle Eintraege konform oder gefiltert." "OK"
    exit 0
}

Show-ActionPlan -Plan $toShow

Write-Host "  Zusammenfassung:" -ForegroundColor White
Write-Host ("  UPDATE     : {0,5}" -f $cntUpdate)  -ForegroundColor Cyan
Write-Host ("  DEAKTIV    : {0,5}" -f $cntDeaktiv) -ForegroundColor Red
Write-Host ("  GRP-SYNC   : {0,5}" -f $cntGrp)     -ForegroundColor Yellow
Write-Host ("  SKIP       : {0,5}" -f $cntSkip)     -ForegroundColor DarkGray
Write-Host ""

if ($WhatIfPreference) {
    Write-Host "  [TROCKENLAUF] Keine Aenderungen werden geschrieben." -ForegroundColor Yellow
    Write-Host ""
}

$confirm = Read-Host "  Aenderungen jetzt anwenden? (j/n)"
if ($confirm -ne 'j') {
    Write-Log "Abgebrochen durch Benutzer." "WARN"
    exit 0
}

# ── AUSFUEHREN ───────────────────────────────────────────────────
Write-Section "AUSFUEHREN"
$t0       = $SW.Elapsed.TotalSeconds
$ok       = 0
$err      = 0
$skip     = 0
$idx      = 0
$total    = $toShow.Count

foreach ($entry in $ActionPlan) {
    if ($entry.Aktion -eq "SKIP") { $skip++; continue }

    $idx++
    $lid = $entry.LID
    $row = $entry.Row

    # Fortschritt
    $pct  = [math]::Round(($idx / $total) * 100)
    $fill = [math]::Round(30 * $pct / 100)
    $bar  = ("#" * $fill) + ("." * (30-$fill))
    Write-Host ("`r  [{0}]  {1,4}/{2}  ({3,3}%)  {4,-16}" -f $bar,$idx,$total,$pct,$lid) `
        -NoNewline -ForegroundColor Cyan
    Write-Progress -Activity "Anwenden" -Status "$idx / $total  ($pct%)" -PercentComplete $pct

    $logEntry = [ordered]@{
        LID           = $lid
        Aktion        = $entry.Aktion
        Codes         = $entry.Codes
        Loeschen      = $entry.Loeschen
        Status        = ""
        Details       = ""
        Timestamp     = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
    }

    try {
        # AD-Objekt holen
        $adUser = Get-ADUser -Filter "SamAccountName -eq '$lid'" `
                      -Properties GivenName,Surname,DisplayName,Description,City,
                                  Office,Department,info,EmailAddress,DistinguishedName `
                      -ErrorAction SilentlyContinue

        if (-not $adUser -and $entry.Aktion -ne "DEAKTIV") {
            Write-Log "$lid : AD-Objekt nicht gefunden -> SKIP" "WARN"
            $logEntry.Status  = "SKIP"
            $logEntry.Details = "AD-Objekt nicht gefunden"
            $skip++
            [void]$ResultLog.Add([PSCustomObject]$logEntry)
            continue
        }

        # ── DEAKTIVIEREN ──────────────────────────────────────────
        if ($entry.LoeschenFlag) {
            if ($WhatIfPreference) {
                Write-Log "[DRY] $lid : wuerde deaktiviert" "DRYRUN"
            } else {
                if ($PSCmdlet.ShouldProcess($lid, "Disable-ADAccount")) {
                    Disable-ADAccount -Identity $lid -ErrorAction Stop
                    Write-Log "$lid : Konto deaktiviert" "OK"
                    Send-ChangeNotification -LID $lid -Email (Get-Str $adUser.EmailAddress) `
                        -Changes "Konto deaktiviert (LOESCHEN=$($entry.Loeschen))" -ActionType "Deaktiv"
                }
            }
            $logEntry.Status  = "OK-DEAKTIV"
            $logEntry.Details = "Konto deaktiviert"
            $ok++
        }

        # ── PROPERTY UPDATE ───────────────────────────────────────
        if ($entry.Codes -ne "" -and $adUser) {
            $codes      = $entry.Codes -split ',' | Where-Object { $_ -ne "" }
            $setParams  = @{ Identity = $lid; ErrorAction = "Stop" }
            $otherAttrs = @{}
            $changeSummary = New-Object 'System.Collections.Generic.List[string]'

            foreach ($code in $codes) {
                if (-not $CodeMap.ContainsKey($code)) {
                    Write-Log "$lid : Unbekannter Code '$code' -> ignoriert" "WARN"
                    continue
                }
                $mapping = $CodeMap[$code]
                $newVal  = Get-Str $row.($mapping.CsvCol)
                $adAttr  = $mapping.AdAttr

                if ([string]::IsNullOrEmpty($newVal)) {
                    Write-Log "$lid : Code $code -> Zielwert leer -> uebersprungen" "WARN"
                    continue
                }

                Write-Log "$lid : $code  $adAttr  =  '$newVal'" "DBG"
                $changeSummary.Add("$code=$newVal")

                # 'info' ist kein Standard Set-ADUser Parameter -> OtherAttributes
                if ($adAttr -eq "OtherAttributes") {
                    $otherAttrs["info"] = $newVal
                } else {
                    $setParams[$adAttr] = $newVal
                }
            }

            if ($otherAttrs.Count -gt 0) {
                $setParams["OtherAttributes"] = $otherAttrs
            }

            if ($setParams.Count -gt 2) {   # mehr als nur Identity + ErrorAction
                if ($WhatIfPreference) {
                    Write-Log "[DRY] $lid : wuerde gesetzt: $($changeSummary -join '  |  ')" "DRYRUN"
                } else {
                    if ($PSCmdlet.ShouldProcess($lid, "Set-ADUser ($($codes -join ','))")) {
                        Set-ADUser @setParams
                        Write-Log "$lid : Properties aktualisiert [$($codes -join ',')]" "OK"
                        Send-ChangeNotification -LID $lid -Email (Get-Str $adUser.EmailAddress) `
                            -Changes ($changeSummary -join " | ") -ActionType "Update"
                    }
                }
                if ($logEntry.Status -eq "") { $logEntry.Status = "OK-UPDATE" }
                $logEntry.Details += $changeSummary -join " | "
                if (-not $entry.LoeschenFlag) { $ok++ }
            } else {
                Write-Log "$lid : Alle Zielwerte leer -> kein Set-ADUser" "SKIP"
                if ($logEntry.Status -eq "") { $logEntry.Status = "SKIP-LEER" }
                $skip++
            }
        }

        # ── GRUPPEN SYNC ──────────────────────────────────────────
        if ($GroupSync -and $adUser) {
            $grpChanges = 0

            foreach ($grp in $entry.GrpAdd) {
                try {
                    $grpObj = Get-ADGroup -Filter "Name -eq '$grp'" -ErrorAction SilentlyContinue
                    if ($grpObj) {
                        $isMember = (Get-ADGroupMember $grpObj -ErrorAction SilentlyContinue |
                                     Where-Object { $_.SamAccountName -eq $lid }).Count -gt 0
                        if (-not $isMember) {
                            if ($WhatIfPreference) {
                                Write-Log "[DRY] $lid : wuerde zu Gruppe '$grp' hinzugefuegt" "DRYRUN"
                            } else {
                                if ($PSCmdlet.ShouldProcess($lid, "Add-ADGroupMember $grp")) {
                                    Add-ADGroupMember -Identity $grpObj -Members $lid -ErrorAction Stop
                                    Write-Log "$lid : + Gruppe '$grp'" "OK"
                                }
                            }
                            $grpChanges++
                        }
                    } else {
                        Write-Log "$lid : Gruppe '$grp' nicht gefunden -> uebersprungen" "WARN"
                    }
                } catch {
                    Write-Log "$lid : Fehler Gruppe hinzufuegen '$grp': $_" "ERR"
                }
            }

            foreach ($grp in $entry.GrpRemove) {
                try {
                    $grpObj = Get-ADGroup -Filter "Name -eq '$grp'" -ErrorAction SilentlyContinue
                    if ($grpObj) {
                        $isMember = (Get-ADGroupMember $grpObj -ErrorAction SilentlyContinue |
                                     Where-Object { $_.SamAccountName -eq $lid }).Count -gt 0
                        if ($isMember) {
                            if ($WhatIfPreference) {
                                Write-Log "[DRY] $lid : wuerde aus Gruppe '$grp' entfernt" "DRYRUN"
                            } else {
                                if ($PSCmdlet.ShouldProcess($lid, "Remove-ADGroupMember $grp")) {
                                    Remove-ADGroupMember -Identity $grpObj -Members $lid -Confirm:$false -ErrorAction Stop
                                    Write-Log "$lid : - Gruppe '$grp'" "OK"
                                }
                            }
                            $grpChanges++
                        }
                    }
                } catch {
                    Write-Log "$lid : Fehler Gruppe entfernen '$grp': $_" "ERR"
                }
            }

            if ($grpChanges -gt 0) {
                $logEntry.Details += " | GRP-SYNC:$grpChanges"
                if ($logEntry.Status -eq "") {
                    $logEntry.Status = "OK-GRP"
                    $ok++
                }
            }
        }

        if ($logEntry.Status -eq "") {
            $logEntry.Status = "SKIP-NOCHANGE"
            $skip++
        }

    } catch {
        $err++
        $errMsg = $_.Exception.Message
        Write-Log "$lid : FEHLER: $errMsg" "ERR"
        $logEntry.Status  = "FEHLER"
        $logEntry.Details = $errMsg
    }

    [void]$ResultLog.Add([PSCustomObject]$logEntry)
}

Write-Host ""
Write-Progress -Activity "Anwenden" -Completed

# ── ERGEBNIS-CSV ─────────────────────────────────────────────────
Write-Section "EXPORT ERGEBNIS-LOG"
$LogCsvPath = Join-Path $ScriptDir "Apply_Result_$Timestamp.csv"
$ResultLog | Export-Csv -Path $LogCsvPath -Delimiter ';' -NoTypeInformation -Encoding UTF8
Write-Log "Ergebnis-CSV: $LogCsvPath" "OK"

# ── ABSCHLUSS ────────────────────────────────────────────────────
$SW.Stop()
$dur = $SW.Elapsed.ToString('mm\:ss\.ff')
$mb  = [math]::Round([System.GC]::GetTotalMemory($false) / 1MB, 1)

Write-Host ""
Write-Host "  +----------------------------------------------------------+" -ForegroundColor Green
Write-Host "  |  ABGESCHLOSSEN                                           |" -ForegroundColor Green
Write-Host ("  |  OK (Update/Deaktiv) : {0,-37}|" -f $ok)   -ForegroundColor Green
Write-Host ("  |  Fehler              : {0,-37}|" -f $err)   -ForegroundColor $(if($err -gt 0){"Red"}else{"Green"})
Write-Host ("  |  Uebersprungen       : {0,-37}|" -f $skip)  -ForegroundColor DarkGreen
Write-Host ("  |  Dauer               : {0,-37}|" -f $dur)   -ForegroundColor DarkGreen
Write-Host ("  |  RAM                 : {0,-37}|" -f "${mb} MB") -ForegroundColor DarkGreen
Write-Host "  |                                                          |" -ForegroundColor DarkGreen
Write-Host ("  |  Result-CSV : {0,-44}|" -f (Split-Path $LogCsvPath -Leaf)) -ForegroundColor DarkGreen
Write-Host ("  |  Log        : {0,-44}|" -f (Split-Path $LogFile -Leaf))    -ForegroundColor DarkGreen
Write-Host "  +----------------------------------------------------------+" -ForegroundColor Green
Write-Host ""

Write-Log ("Fertig. OK={0}  ERR={1}  SKIP={2}  Dauer={3}  WhatIf={4}" -f `
    $ok, $err, $skip, $dur, $WhatIfPreference) "PERF"
