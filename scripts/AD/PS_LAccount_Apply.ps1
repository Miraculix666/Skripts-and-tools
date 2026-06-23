# FileName: PS_LAccount_Apply.ps1
# Version:  2.0
# Beschreibung: Liest die Properties-CSV des AD-Sync-Managers (v9.x) und
#               wendet Aenderungen auf Active Directory an.
#
# EINGABE:  L-Kennungen_Properties_vX.X.csv  (aus PS_LAccount_Manager v9.x)
# AUSGABE:  Apply_Result_<Timestamp>.csv  +  Apply_Log_<Timestamp>.log
#
# AENDERUNGEN v2.0:
#   - E-Mail-Versand entfernt
#   - -ColFilter: Filter nach Spaltennamen der CSV (statt Code-Namen)
#   - Gruppenspaltennamen ohne "GRP_"-Prefix ausgelesen
#   - -GroupFilter: Regex/Substring-Filter auf Gruppennamen (z.B. "test")
#     Wenn der Filter zutrifft, wird "X" gesetzt (case-insensitiv)
#
# AKTIONEN (gesteuert durch Tabellenspalten):
#   NICHT_KONFORM  -> Set-ADUser fuer betroffene Attribute
#   LOESCHEN <> "" -> Konto deaktivieren (kein Hard-Delete)
#   Gruppenspalten -> Gruppen-Sync (X=hinzufuegen, ""=entfernen)
#
# NICHT_KONFORM CODES:
#   VN   = GivenName      (Vorname)        <- Spalte AENDERN_Vorname
#   NN   = Surname        (Nachname)       <- Spalte AENDERN_Nachname
#   DN   = DisplayName                    <- Spalte AENDERN_DisplayName
#   DESC = Description                    <- Spalte AENDERN_Description
#   ORT  = City / l       (Ort)           <- Spalte AENDERN_Ort
#   GEB  = Office         (Buero)         <- Spalte AENDERN_Buero
#   DEZ  = Department                     <- Spalte AENDERN_Dez
#   INFO = info                           <- Spalte AENDERN_Info
#
# PARAMETER:
#   -CsvPath       Pfad zur Properties-CSV  (Pflicht)
#   -ColFilter     Komma-/Semikolonliste von CSV-Spaltennamen, die verarbeitet
#                  werden sollen. Leer = alle NICHT_KONFORM-Spalten.
#                  Spaltennamen exakt wie in der CSV, z.B.:
#                    -ColFilter "AENDERN_Vorname,AENDERN_Nachname,AENDERN_DisplayName"
#                    -ColFilter "LOESCHEN"        (nur Deaktivierungen)
#   -GroupSync     Gruppen-Sync ausfuehren  (Default: nein)
#   -GroupFilter   Substring (case-insensitiv) zum Vorfiltern von Gruppennamen.
#                  Nur Gruppen, deren Name diesen Teilstring enthaelt, werden
#                  verarbeitet. Passende Gruppen bekommen "X" gesetzt.
#                  z.B. -GroupFilter "test"  -> nur *test*-Gruppen werden beruecksichtigt
#   -WhatIf        Trockenlauf: zeigt alles, aendert nichts
#   -DebugMode     Ausfuehrliches Logging
#
# BEISPIELE:
#   # Trockenlauf, alle Korrekturen + Gruppen-Sync:
#   .\PS_LAccount_Apply.ps1 -CsvPath .\Properties.csv -GroupSync -WhatIf
#
#   # Nur Vor-/Nachname und DisplayName anpassen:
#   .\PS_LAccount_Apply.ps1 -CsvPath .\Properties.csv -ColFilter "AENDERN_Vorname,AENDERN_Nachname,AENDERN_DisplayName"
#
#   # Gruppen-Sync, aber nur Gruppen die "Schulung" im Namen haben:
#   .\PS_LAccount_Apply.ps1 -CsvPath .\Properties.csv -GroupSync -GroupFilter "Schulung"
#
#   # Nur Deaktivierungen ausfuehren:
#   .\PS_LAccount_Apply.ps1 -CsvPath .\Properties.csv -ColFilter "LOESCHEN"

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(Mandatory=$true)]
    [string] $CsvPath,

    [Parameter(Mandatory=$false)]
    [string] $ColFilter    = "",     # Spaltennamen-Filter, leer = alles

    [Parameter(Mandatory=$false)]
    [switch] $GroupSync,

    [Parameter(Mandatory=$false)]
    [string] $GroupFilter  = "",     # Substring-Filter fuer Gruppennamen

    [Parameter(Mandatory=$false)]
    [switch] $DebugMode
)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

$Version   = "2.0"
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile   = Join-Path $ScriptDir "Apply_Log_$Timestamp.log"
$SW        = [System.Diagnostics.Stopwatch]::StartNew()

# ══════════════════════════════════════════════════════════════════
#  UI & LOGGING
# ══════════════════════════════════════════════════════════════════
function Show-Banner {
    Clear-Host
    Write-Host ""
    Write-Host "  +----------------------------------------------------------+" -ForegroundColor DarkCyan
    Write-Host "  |  AD L-KENNUNG APPLY MANAGER  v$Version                        |" -ForegroundColor Cyan
    Write-Host "  |  $(Get-Date -Format 'dd.MM.yyyy  HH:mm:ss')                               |" -ForegroundColor DarkCyan
    if ($WhatIfPreference) {
    Write-Host "  |  !! TROCKENLAUF  -  KEINE AD-AENDERUNGEN !!             |" -ForegroundColor Yellow
    }
    if ($ColFilter -ne "") {
    Write-Host ("  |  ColFilter  : {0,-44}|" -f $ColFilter)  -ForegroundColor DarkYellow
    }
    if ($GroupFilter -ne "") {
    Write-Host ("  |  GrpFilter  : '*{0}*'  (case-insensitiv){1}|" -f $GroupFilter, " ".PadRight([math]::Max(0,42-$GroupFilter.Length))) -ForegroundColor DarkYellow
    }
    if ($GroupSync) {
    Write-Host "  |  Gruppen-Sync: aktiv                                     |" -ForegroundColor Cyan
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
        "OK"     { "[OK]   " } "WARN"   { "[WARN] " } "ERR"    { "[ERR]  " }
        "DBG"    { "[DBG]  " } "STEP"   { "[----] " } "SKIP"   { "[SKIP] " }
        "DRYRUN" { "[DRY]  " }
        default  { "[INFO] " }
    }
    $color = switch ($L) {
        "OK"     { "Green"       } "WARN"   { "Yellow"     } "ERR"    { "Red"       }
        "DBG"    { "Magenta"     } "STEP"   { "Cyan"       } "SKIP"   { "DarkGray"  }
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
    Write-Host ("  {0,-16}  {1,-10}  {2,-30}  {3}" -f "L-Kennung","Aktion","Codes","Loeschen") -ForegroundColor DarkCyan
    Write-Host ("  " + ("-" * 76)) -ForegroundColor DarkGray
    foreach ($item in $Plan) {
        $color = switch ($item.Aktion) {
            "UPDATE"   { "Cyan"     }
            "DEAKTIV"  { "Red"      }
            "GRP-SYNC" { "Yellow"   }
            default    { "DarkGray" }
        }
        Write-Host ("  {0,-16}  {1,-10}  {2,-30}  {3}" -f `
            $item.LID, $item.Aktion, $item.Codes, $item.Loeschen) -ForegroundColor $color
    }
    Write-Host ""
}

# ══════════════════════════════════════════════════════════════════
#  CODE -> AD-ATTRIBUT MAPPING
#  Key   = NICHT_KONFORM-Code
#  Value = AD-Parameter fuer Set-ADUser  +  Quellspalte in der CSV
# ══════════════════════════════════════════════════════════════════
$CodeMap = [ordered]@{
    "VN"   = @{ AdAttr = "GivenName";    CsvCol = "AENDERN_Vorname"      }
    "NN"   = @{ AdAttr = "Surname";      CsvCol = "AENDERN_Nachname"     }
    "DN"   = @{ AdAttr = "DisplayName";  CsvCol = "AENDERN_DisplayName"  }
    "DESC" = @{ AdAttr = "Description";  CsvCol = "AENDERN_Description"  }
    "ORT"  = @{ AdAttr = "City";         CsvCol = "AENDERN_Ort"          }
    "GEB"  = @{ AdAttr = "Office";       CsvCol = "AENDERN_Buero"        }
    "DEZ"  = @{ AdAttr = "Department";   CsvCol = "AENDERN_Dez"          }
    "INFO" = @{ AdAttr = "OtherAttr";    CsvCol = "AENDERN_Info"         }
}

# Umgekehrtes Mapping: CsvSpaltenname -> Code  (fuer -ColFilter)
$ColToCode = @{}
foreach ($k in $CodeMap.Keys) { $ColToCode[$CodeMap[$k].CsvCol] = $k }

# ══════════════════════════════════════════════════════════════════
#  HILFSFUNKTIONEN
# ══════════════════════════════════════════════════════════════════
function Get-Str { param($v)
    if ($null -eq $v) { return "" }
    return $v.ToString().Trim()
}

# ══════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════
Show-Banner
Import-Module ActiveDirectory -Verbose:$false

Write-Log "PS v$($PSVersionTable.PSVersion)  PID=$PID  WhatIf=$($WhatIfPreference)" "DBG"

try {
    $null = Get-ADDomain
    Write-Log "AD-Verbindung OK" "OK"
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
    $AllRows = Import-Csv -Path $CsvPath -Delimiter ';' -Encoding Default
    Write-Log "Fallback auf Encoding Default" "WARN"
}

Write-Log ("CSV: {0} Zeilen geladen  [{1:F2}s]" -f $AllRows.Count, ($SW.Elapsed.TotalSeconds-$t0)) "OK"

# Spaltenstruktur analysieren
$AllCols    = $AllRows[0].psobject.Properties.Name

# Gruppen-Spalten: Tabelle kann "GRP_Gruppenname" ODER direkt "Gruppenname" enthalten.
# Beide Varianten werden unterstuetzt. Resultat: immer der reine Gruppenname.
$RawGrpCols = @($AllCols | Where-Object { $_ -like "GRP_*" -or $_ -eq "Gruppen" })
$GrpCols    = @()   # Hashtable: @{ ColName; GrpName }

foreach ($col in $AllCols) {
    if ($col -like "GRP_*") {
        $grpName = $col -replace '^GRP_', ''

        # GroupFilter anwenden (case-insensitiv Substring)
        if ($GroupFilter -ne "" -and $grpName -notlike "*$GroupFilter*") {
            Write-Log "GrpFilter: '$grpName' uebersprungen (kein '*$GroupFilter*')" "DBG"
            continue
        }

        $GrpCols += [PSCustomObject]@{ ColName = $col; GrpName = $grpName }

    } elseif ($col -eq "Gruppen") {
        # GroupMode=Single: alle Gruppen in einer Zelle, semikolonsepariert
        # -> wird spaeter zeilenweise aufgeloest
        $GrpCols += [PSCustomObject]@{ ColName = "Gruppen"; GrpName = "__SINGLE__" }
    }
}

Write-Log ("Spalten: {0} gesamt  |  {1} Gruppen-Spalten nach Filter" -f $AllCols.Count, $GrpCols.Count) "DBG"

# ColFilter aufloesen: Spaltennamen -> Codes
# Leer = alle Codes aktiv. "LOESCHEN" als Sonderfall direkt auswertbar.
$ActiveCodes    = @()
$OnlyLoeschen   = $false

if ($ColFilter -ne "") {
    $filterCols = $ColFilter -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    Write-Log "ColFilter Spalten: $($filterCols -join '  |  ')" "WARN"

    foreach ($fc in $filterCols) {
        if ($fc -ieq "LOESCHEN") {
            $OnlyLoeschen = $true
        } elseif ($ColToCode.ContainsKey($fc)) {
            $ActiveCodes += $ColToCode[$fc]
        } else {
            Write-Log "ColFilter: Spalte '$fc' nicht im Mapping bekannt -> ignoriert" "WARN"
        }
    }
    Write-Log ("Aktive Codes aus ColFilter: {0}  |  OnlyLoeschen={1}" -f `
        ($ActiveCodes -join ','), $OnlyLoeschen) "DBG"
} else {
    # Kein Filter -> alle Codes aktiv
    $ActiveCodes = @($CodeMap.Keys)
    Write-Log "ColFilter: keiner -> alle Codes aktiv" "DBG"
}

# ── AKTIONSPLAN AUFBAUEN ─────────────────────────────────────────
Write-Section "AKTIONSPLAN"
$ActionPlan = New-Object 'System.Collections.Generic.List[PSObject]'
$ResultLog  = New-Object 'System.Collections.Generic.List[PSObject]'

$cntUpdate  = 0
$cntDeaktiv = 0
$cntGrp     = 0
$cntSkip    = 0

foreach ($row in $AllRows) {
    $lid       = Get-Str $row."L-Kennung"
    $nkCodes   = Get-Str $row."NICHT_KONFORM"
    $loeschen  = Get-Str $row."LOESCHEN"
    $geloescht = Get-Str $row."GELOESCHT"

    if ([string]::IsNullOrEmpty($lid)) { continue }

    # ── Gruppen ermitteln ──
    $grpAdd    = New-Object 'System.Collections.Generic.List[string]'
    $grpRemove = New-Object 'System.Collections.Generic.List[string]'

    if ($GroupSync) {
        foreach ($gc in $GrpCols) {
            if ($gc.GrpName -eq "__SINGLE__") {
                # GroupMode=Single: Zelle "Gruppen" splitten
                $cellVal = Get-Str $row."Gruppen"
                $names   = $cellVal -split ';' | Where-Object { $_ -ne "" }
                foreach ($n in $names) {
                    # GroupFilter auf Einzelnamen anwenden
                    if ($GroupFilter -ne "" -and $n -notlike "*$GroupFilter*") { continue }
                    $grpAdd.Add($n)
                }
            } else {
                $val = Get-Str $row.($gc.ColName)
                # GroupFilter wurde schon beim Aufbau von $GrpCols angewendet
                if ($val -ieq "X") { $grpAdd.Add($gc.GrpName) }
                else               { $grpRemove.Add($gc.GrpName) }
            }
        }
    }

    # ── Aktionstyp bestimmen ──
    $aktion       = "SKIP"
    $aktCodes     = @()
    $loeschenFlag = $false

    # 1. Deaktivieren?
    if ($loeschen -ne "" -and (-not $OnlyLoeschen -or $ColFilter -match "LOESCHEN")) {
        # Nur wenn kein reiner ColFilter auf andere Spalten gesetzt ist,
        # ODER wenn LOESCHEN explizit in ColFilter enthalten ist
        if ($ColFilter -eq "" -or $OnlyLoeschen -or ($filterCols -icontains "LOESCHEN")) {
            $loeschenFlag = $true
            $aktion       = "DEAKTIV"
            $cntDeaktiv++
        }
    }

    # 2. Property-Update?
    if ($nkCodes -ne "" -and $geloescht -ne "XXX" -and -not $OnlyLoeschen) {
        $rawCodes = $nkCodes -split '[,;\s]' | Where-Object { $_ -ne "" }
        $aktCodes = @($rawCodes | Where-Object { $ActiveCodes -icontains $_ })

        if ($aktCodes.Count -gt 0) {
            if ($aktion -eq "SKIP") { $cntUpdate++ }
            $aktion = if ($aktion -eq "DEAKTIV") { "DEAKTIV+UPD" } else { "UPDATE" }
        }
    }

    # 3. Gruppen-Sync?
    if ($GroupSync -and ($grpAdd.Count -gt 0 -or $grpRemove.Count -gt 0)) {
        if ($aktion -eq "SKIP") { $cntGrp++ }
        $aktion = if ($aktion -eq "SKIP") { "GRP-SYNC" } else { $aktion + "+GRP" }
    }

    if ($aktion -eq "SKIP") { $cntSkip++ }

    [void]$ActionPlan.Add([PSCustomObject]@{
        LID          = $lid
        Aktion       = $aktion
        Codes        = ($aktCodes -join ",")
        Loeschen     = $loeschen
        GrpAdd       = @($grpAdd)
        GrpRemove    = @($grpRemove)
        Row          = $row
        LoeschenFlag = $loeschenFlag
    })
}

Write-Log ("Plan: {0} UPDATE  {1} DEAKTIV  {2} GRP-SYNC  {3} SKIP" -f `
    $cntUpdate, $cntDeaktiv, $cntGrp, $cntSkip) "INFO"

$toShow = @($ActionPlan | Where-Object { $_.Aktion -ne "SKIP" })
if ($toShow.Count -eq 0) {
    Write-Log "Keine Aktionen notwendig. Alle Eintraege konform oder durch Filter ausgeschlossen." "OK"
    exit 0
}

Show-ActionPlan -Plan $toShow

Write-Host ("  {0,-12}: {1,5}" -f "UPDATE",  $cntUpdate)  -ForegroundColor Cyan
Write-Host ("  {0,-12}: {1,5}" -f "DEAKTIV", $cntDeaktiv) -ForegroundColor Red
Write-Host ("  {0,-12}: {1,5}" -f "GRP-SYNC",$cntGrp)     -ForegroundColor Yellow
Write-Host ("  {0,-12}: {1,5}" -f "SKIP",    $cntSkip)     -ForegroundColor DarkGray
Write-Host ""

if ($WhatIfPreference) {
    Write-Host "  [TROCKENLAUF] Keine Aenderungen werden vorgenommen." -ForegroundColor Yellow
    Write-Host ""
}

$confirm = Read-Host "  Aenderungen jetzt anwenden? (j/n)"
if ($confirm -ine 'j') {
    Write-Log "Abgebrochen durch Benutzer." "WARN"; exit 0
}

# ── AUSFUEHREN ───────────────────────────────────────────────────
Write-Section "AUSFUEHREN"
$t0    = $SW.Elapsed.TotalSeconds
$ok    = 0
$err   = 0
$skip  = 0
$idx   = 0
$total = $toShow.Count

foreach ($entry in $ActionPlan) {
    if ($entry.Aktion -eq "SKIP") { $skip++; continue }

    $idx++
    $lid = $entry.LID
    $row = $entry.Row

    $pct  = [math]::Round(($idx / $total) * 100)
    $fill = [math]::Round(30 * $pct / 100)
    $bar  = ("#" * $fill) + ("." * (30 - $fill))
    Write-Host ("`r  [{0}]  {1,4}/{2}  ({3,3}%)  {4,-16}" -f $bar,$idx,$total,$pct,$lid) `
        -NoNewline -ForegroundColor Cyan
    Write-Progress -Activity "Anwenden" -Status "$idx / $total  ($pct%)" -PercentComplete $pct

    $logEntry = [ordered]@{
        LID       = $lid
        Aktion    = $entry.Aktion
        Codes     = $entry.Codes
        Loeschen  = $entry.Loeschen
        Status    = ""
        Details   = ""
        Timestamp = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
    }

    try {
        $adUser = Get-ADUser -Filter "SamAccountName -eq '$lid'" `
                     -Properties GivenName,Surname,DisplayName,Description,`
                                 City,Office,Department,info,DistinguishedName `
                     -ErrorAction SilentlyContinue

        if (-not $adUser) {
            Write-Log "$lid : nicht in AD gefunden -> SKIP" "WARN"
            $logEntry.Status  = "SKIP-NICHT-GEFUNDEN"
            $logEntry.Details = "AD-Objekt nicht gefunden"
            $skip++
            [void]$ResultLog.Add([PSCustomObject]$logEntry)
            continue
        }

        # ── DEAKTIVIEREN ──────────────────────────────────────────
        if ($entry.LoeschenFlag) {
            if ($WhatIfPreference) {
                Write-Log "[DRY] $lid : wuerde deaktiviert  (LOESCHEN=$($entry.Loeschen))" "DRYRUN"
            } else {
                if ($PSCmdlet.ShouldProcess($lid, "Disable-ADAccount")) {
                    Disable-ADAccount -Identity $lid -ErrorAction Stop
                    Write-Log "$lid : deaktiviert" "OK"
                }
            }
            $logEntry.Status  = "OK-DEAKTIV"
            $logEntry.Details = "Deaktiviert (Grund=$($entry.Loeschen))"
            $ok++
        }

        # ── PROPERTY UPDATE ───────────────────────────────────────
        if ($entry.Codes -ne "") {
            $codes         = $entry.Codes -split ',' | Where-Object { $_ -ne "" }
            $setParams     = @{ Identity = $lid; ErrorAction = "Stop" }
            $otherAttrs    = @{}
            $changeSummary = New-Object 'System.Collections.Generic.List[string]'

            foreach ($code in $codes) {
                if (-not $CodeMap.ContainsKey($code)) {
                    Write-Log "$lid : Code '$code' unbekannt -> ignoriert" "WARN"
                    continue
                }
                $m      = $CodeMap[$code]
                $newVal = Get-Str $row.($m.CsvCol)

                if ([string]::IsNullOrEmpty($newVal)) {
                    Write-Log "$lid : $code -> Zielwert leer (Spalte $($m.CsvCol)) -> uebersprungen" "WARN"
                    continue
                }

                Write-Log "$lid : $code  [$($m.AdAttr)]  =  '$newVal'" "DBG"
                $changeSummary.Add("$code=$newVal")

                if ($m.AdAttr -eq "OtherAttr") {
                    # 'info' ist kein direkter Set-ADUser Parameter
                    $otherAttrs["info"] = $newVal
                } else {
                    $setParams[$m.AdAttr] = $newVal
                }
            }

            if ($otherAttrs.Count -gt 0) { $setParams["OtherAttributes"] = $otherAttrs }

            # Nur Set-ADUser aufrufen wenn es mehr als Identity+ErrorAction gibt
            if ($setParams.Count -gt 2) {
                if ($WhatIfPreference) {
                    Write-Log "[DRY] $lid : Set-ADUser  $($changeSummary -join '  |  ')" "DRYRUN"
                } else {
                    if ($PSCmdlet.ShouldProcess($lid, "Set-ADUser [$($entry.Codes)]")) {
                        Set-ADUser @setParams
                        Write-Log "$lid : Properties gesetzt  [$($entry.Codes)]" "OK"
                    }
                }
                if ($logEntry.Status -eq "") {
                    $logEntry.Status = "OK-UPDATE"
                    $ok++
                }
                $logEntry.Details += $changeSummary -join " | "
            } else {
                Write-Log "$lid : Keine gueltigen Zielwerte -> kein Set-ADUser" "SKIP"
                if ($logEntry.Status -eq "") {
                    $logEntry.Status = "SKIP-WERTE-LEER"
                    $skip++
                }
            }
        }

        # ── GRUPPEN SYNC ──────────────────────────────────────────
        if ($GroupSync) {
            $grpOk      = 0
            $grpErr     = 0
            $grpDetails = New-Object 'System.Collections.Generic.List[string]'

            # Hinzufuegen
            foreach ($grp in $entry.GrpAdd) {
                try {
                    $grpObj = Get-ADGroup -Filter "Name -eq '$grp'" -ErrorAction SilentlyContinue
                    if (-not $grpObj) {
                        Write-Log "$lid : Gruppe '$grp' nicht gefunden -> SKIP" "WARN"
                        continue
                    }
                    $isMember = [bool](Get-ADGroupMember $grpObj -ErrorAction SilentlyContinue |
                                       Where-Object { $_.SamAccountName -eq $lid })
                    if (-not $isMember) {
                        if ($WhatIfPreference) {
                            Write-Log "[DRY] $lid : + '$grp'" "DRYRUN"
                        } else {
                            if ($PSCmdlet.ShouldProcess($lid, "Add-ADGroupMember: $grp")) {
                                Add-ADGroupMember -Identity $grpObj -Members $lid -ErrorAction Stop
                                Write-Log "$lid : + Gruppe '$grp'" "OK"
                            }
                        }
                        $grpDetails.Add("+$grp")
                        $grpOk++
                    } else {
                        Write-Log "$lid : '$grp' bereits Mitglied" "DBG"
                    }
                } catch {
                    $grpErr++
                    Write-Log "$lid : Fehler + '$grp': $_" "ERR"
                }
            }

            # Entfernen
            foreach ($grp in $entry.GrpRemove) {
                try {
                    $grpObj = Get-ADGroup -Filter "Name -eq '$grp'" -ErrorAction SilentlyContinue
                    if (-not $grpObj) { continue }
                    $isMember = [bool](Get-ADGroupMember $grpObj -ErrorAction SilentlyContinue |
                                       Where-Object { $_.SamAccountName -eq $lid })
                    if ($isMember) {
                        if ($WhatIfPreference) {
                            Write-Log "[DRY] $lid : - '$grp'" "DRYRUN"
                        } else {
                            if ($PSCmdlet.ShouldProcess($lid, "Remove-ADGroupMember: $grp")) {
                                Remove-ADGroupMember -Identity $grpObj -Members $lid `
                                    -Confirm:$false -ErrorAction Stop
                                Write-Log "$lid : - Gruppe '$grp'" "OK"
                            }
                        }
                        $grpDetails.Add("-$grp")
                        $grpOk++
                    }
                } catch {
                    $grpErr++
                    Write-Log "$lid : Fehler - '$grp': $_" "ERR"
                }
            }

            if ($grpOk -gt 0 -or $grpErr -gt 0) {
                $logEntry.Details += " | GRP: $($grpDetails -join ' ')  (Err=$grpErr)"
                if ($logEntry.Status -eq "") {
                    $logEntry.Status = "OK-GRP"
                    $ok++
                }
            }
        }

        if ($logEntry.Status -eq "") {
            $logEntry.Status = "SKIP-KEINE-AENDERUNG"
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

# ── ERGEBNIS EXPORTIEREN ─────────────────────────────────────────
Write-Section "ERGEBNIS EXPORT"
$ResultCsv = Join-Path $ScriptDir "Apply_Result_$Timestamp.csv"
$ResultLog | Export-Csv -Path $ResultCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
Write-Log "Ergebnis-CSV: $ResultCsv" "OK"

# ── ABSCHLUSS ────────────────────────────────────────────────────
$SW.Stop()
$dur = $SW.Elapsed.ToString('mm\:ss\.ff')
$mb  = [math]::Round([System.GC]::GetTotalMemory($false) / 1MB, 1)

Write-Host ""
Write-Host "  +----------------------------------------------------------+" -ForegroundColor Green
Write-Host "  |  ABGESCHLOSSEN                                           |" -ForegroundColor Green
Write-Host ("  |  OK (Update/Deaktiv)  : {0,-34}|" -f $ok)   -ForegroundColor Green
Write-Host ("  |  Fehler               : {0,-34}|" -f $err)   -ForegroundColor $(if($err -gt 0){"Red"}else{"Green"})
Write-Host ("  |  Uebersprungen        : {0,-34}|" -f $skip)  -ForegroundColor DarkGreen
Write-Host ("  |  Dauer                : {0,-34}|" -f $dur)   -ForegroundColor DarkGreen
Write-Host ("  |  RAM                  : {0,-34}|" -f "${mb} MB") -ForegroundColor DarkGreen
Write-Host "  |                                                          |" -ForegroundColor DarkGreen
Write-Host ("  |  Result : {0,-48}|" -f (Split-Path $ResultCsv -Leaf)) -ForegroundColor DarkGreen
Write-Host ("  |  Log    : {0,-48}|" -f (Split-Path $LogFile -Leaf))    -ForegroundColor DarkGreen
Write-Host "  +----------------------------------------------------------+" -ForegroundColor Green
Write-Host ""

Write-Log ("Fertig. OK={0}  ERR={1}  SKIP={2}  Dauer={3}  WhatIf={4}" -f `
    $ok, $err, $skip, $dur, $WhatIfPreference) "PERF"
