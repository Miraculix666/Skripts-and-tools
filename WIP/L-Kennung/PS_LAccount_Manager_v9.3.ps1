# FileName: PS_LAccount_Manager.ps1
# Description: AD-Sync Tool v9.3. Vollständig robuster ScriptBlock für Runspaces.
# Version: 9.3
# Author: PS-Coding
#
# CHANGELOG v9.3:
#   [FIX-7] KRITISCH: Alle inline 'if'-Ausdrücke im Runspace-ScriptBlock entfernt.
#           Ursache: 'return if (...)' sowie komplexe inline-if in Hashtable-Werten
#           werden in isolierten Runspaces (InitialSessionState.CreateDefault) unter
#           bestimmten PS-Versionen nicht korrekt geparst -> "if nicht erkannt".
#           Fix: ALLE bedingten Werte werden VOR dem Hashtable in eigene Variablen
#           berechnet. Der Hashtable enthält nur noch reine Variablen-Referenzen.
#   [FIX-1] Funktion von 'Start-Process' -> 'Invoke-Main' (eingebautes PS-Cmdlet).
#   [FIX-2] PSDataCollection-Unrolling korrigiert (IsOpen/Count in CSV-Bug).
#   [FIX-3] Doppelte $Anmerkung-Zuweisung korrigiert.
#   [FIX-4] Collection-Modifikation während foreach-Iteration gesichert.
#   [FIX-5] 'Geändert'-Spalte wird jetzt korrekt mit Change-Codes befüllt.
#   [FIX-6] Bounds-Check für $Results[0] hinzugefügt.
#   [OPT-1] Runspace Fehler-Stream wird geloggt.
#   [OPT-2] Write-Progress für Fortschrittsanzeige.
#   [OPT-3] Robustes CSV-Escaping für Semikolon/Anführungszeichen in Werten.

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)] [string]$CsvPath,
    [Parameter(Mandatory=$false)] [string]$RequiredCsvPath,
    [Parameter(Mandatory=$false)] [switch]$SearchGlobal,
    [Parameter(Mandatory=$false)] [int]$TestCount   = 0,
    [Parameter(Mandatory=$false)] [int]$MaxThreads  = 8,
    [Parameter(Mandatory=$false)] [switch]$DebugMode
)

$Version              = "9.3"
$ErrorActionPreference = 'Stop'
$ScriptDir            = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD }
$Timestamp            = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile              = Join-Path $ScriptDir "PS_LAccountManager_$Timestamp.log"
$Stopwatch            = [System.Diagnostics.Stopwatch]::StartNew()

function Show-Header {
    Clear-Host
    Write-Host @"
██╗      █████╗  ██████╗ ██████╗ ██████╗ ██╗███╗   ██╗ ██████╗
██║     ██╔══██╗██╔════╝██╔════╝██╔═══██╗██║████╗  ██║██╔════╝
██║     ███████║██║     ██║     ██║   ██║██║██╔██╗ ██║██║  ███╗
██║     ██╔══██║██║     ██║     ██║   ██║██║██║╚██╗██║██║   ██║
███████╗██║  ██║╚██████╗╚██████╗╚██████╔╝██║██║ ╚████║╚██████╔╝
              AD COMPLIANCE & SYNC MANAGER v$Version (MT)
----------------------------------------------------------------
STATUS: Live-Reporting Aktiv | Threads: $MaxThreads
"@ -ForegroundColor Cyan
}

function Write-Log {
    param([string]$Message,
          [ValidateSet("INFO","WARN","ERROR","SUCCESS","DEBUG")] $Level = "INFO")
    if ($Level -eq "DEBUG" -and -not $DebugMode) { return }
    $Time     = Get-Date -Format "HH:mm:ss"
    $LogColor = switch ($Level) {
        "SUCCESS" { "Green"   }
        "ERROR"   { "Red"     }
        "WARN"    { "Yellow"  }
        "DEBUG"   { "Magenta" }
        default   { "Gray"    }
    }
    $Msg = "[$Time] [$Level] $Message"
    Write-Host $Msg -ForegroundColor $LogColor
    $Msg | Out-File -FilePath $LogFile -Append
}


function Get-RequiredList {
    param([string]$RequiredCsvPath)
    # 1. BEDARFSLISTE
    $RequiredList = New-Object System.Collections.Generic.HashSet[string]
    if ($RequiredCsvPath -and (Test-Path $RequiredCsvPath)) {
        Write-Log "Lade Bedarfsliste..." "DEBUG"
        foreach ($line in (Get-Content $RequiredCsvPath -Encoding UTF8)) {
            if ($line -match "(L\d{6,8})") { [void]$RequiredList.Add($matches[1].ToUpper()) }
        }
        Write-Log "Bedarfsliste: $($RequiredList.Count) Eintraege." "SUCCESS"
    }

    return $RequiredList
}

function Get-MasterCsvData {
    param([string]$CsvPath)
    # 2. MASTER-CSV
    Write-Log "Lese Master-CSV..." "DEBUG"
    $MasterCsvData = @{}
    foreach ($line in (Import-Csv -Path $CsvPath -Delimiter ';' -Encoding Default)) {
        $key = if ($line."L-Kennung") { $line."L-Kennung".ToString().Trim().ToUpper() } else { "" }
        if ($key) { $MasterCsvData[$key] = $line }
    }
    Write-Log "Master-CSV: $($MasterCsvData.Count) Zeilen." "SUCCESS"

    return $MasterCsvData
}

function Get-ADData {
    param([switch]$SearchGlobal, [hashtable]$MasterCsvData, [int]$TestCount)
    # 3. AD-DISCOVERY
    Write-Log "Schritt 1: AD-Discovery (OUs 81/82)..." "INFO"
    $ADCache      = @{}
    $UniqueGroups = New-Object System.Collections.Generic.HashSet[string]

    $TargetOUs = Get-ADOrganizationalUnit -Filter "Name -eq '81' -or Name -eq '82'" `
                     -ErrorAction SilentlyContinue
    foreach ($ou in $TargetOUs) {
        Write-Log "Scanne OU: $($ou.DistinguishedName)" "DEBUG"
        $users = Get-ADUser -Filter * -SearchBase $ou.DistinguishedName `
                     -Properties DisplayName,Description,GivenName,Surname,l,
                                  physicalDeliveryOfficeName,department,info,MemberOf
        foreach ($u in $users) {
            $ADCache[$u.SamAccountName.ToUpper()] = $u
            foreach ($g in $u.MemberOf) {
                # FIX-8: Komma in -replace 'x','y' wird als 2. Methodenargument geparst.
                # Loesung: Ausdruck vorab in Variable berechnen.
                $gn = ($g -split ',')[0] -replace 'CN=',''
                [void]$UniqueGroups.Add($gn)
            }
        }
    }

    if ($SearchGlobal) {
        Write-Log "Schritt 2: Globaler AD-Check..." "INFO"
        $GlobalUsers = Get-ADUser -Filter "SamAccountName -like 'L*'" `
                           -Properties DisplayName,Description,GivenName,Surname,l,
                                        physicalDeliveryOfficeName,department,info,MemberOf
        foreach ($gu in $GlobalUsers) {
            $sam = $gu.SamAccountName.ToUpper()
            if (-not $ADCache.ContainsKey($sam)) {
                $ADCache[$sam] = $gu
                foreach ($g in $gu.MemberOf) {
                    $gn = ($g -split ',')[0] -replace 'CN=',''
                    [void]$UniqueGroups.Add($gn)
                }
            }
        }
    }

    $SortedGroups   = $UniqueGroups | Sort-Object
    $AllUniqueSAMs  = $ADCache.Keys + $MasterCsvData.Keys | Select-Object -Unique | Sort-Object
    $ProcessList    = if ($TestCount -gt 0) {
        Write-Log "TESTMODUS: Begrenze auf $TestCount Eintraege." "WARN"
        $AllUniqueSAMs | Select-Object -First $TestCount
    } else { $AllUniqueSAMs }

    Write-Log ("Discovery beendet ({0}s). Starte Verarbeitung von {1} Eintraegen..." `
               -f $Stopwatch.Elapsed.TotalSeconds.ToString('F2'), $ProcessList.Count) "SUCCESS"

    return @{
        ADCache = $ADCache
        UniqueGroups = $UniqueGroups
        SortedGroups = $SortedGroups
        AllUniqueSAMs = $AllUniqueSAMs
        ProcessList = $ProcessList
    }
}

function Process-Runspaces {
    param(
        [int]$MaxThreads,
        [array]$ProcessList,
        [hashtable]$MasterCsvData,
        [hashtable]$ADCache,
        [System.Collections.Generic.HashSet[string]]$RequiredList,
        [array]$SortedGroups,
        [string]$LogFile
    )
    # 4. RUNSPACE-POOL
    $Pool = [runspacefactory]::CreateRunspacePool(
        1, $MaxThreads,
        [system.management.automation.runspaces.initialsessionstate]::CreateDefault(),
        $Host)
    $Pool.Open()

    # -----------------------------------------------------------------------
    # SCRIPTBLOCK: Alle bedingten Werte werden VOR dem Hashtable in Variablen
    # berechnet. KEINE inline-if-Ausdruecke im Hashtable. (FIX-7)
    # -----------------------------------------------------------------------
    $ScriptBlock = {
        param($LID, $CsvRow, $ADObj, $RequiredList, $SortedGroups)

        # --- Hilfsfunktionen ---
        function IsDiff($a, $b) {
            $x = if ($null -ne $a) { $a.ToString().Trim() } else { "" }
            $y = if ($null -ne $b) { $b.ToString().Trim() } else { "" }
            return ($x -ine $y)
        }

        function GetVerfahren($Row) {
            if ($null -eq $Row) { return "[alle_Verfahren]" }
            $cols   = "Viva","Findus","MobiApps","AccVisio","Verfahren5",
                      "Verfahren6","Verfahren7","Verfahren8","Verfahren9","Verfahren10"
            $active = foreach ($c in $cols) {
                if ($Row.$c -and $Row.$c.ToString().Trim().ToLower() -eq "x") { $c }
            }
            # Kein 'return if' - Ergebnis in Variable, dann return
            $verfahrenStr = if ($active) { $active -join " - " } else { "[alle_Verfahren]" }
            return $verfahrenStr
        }

        # --- Basiswerte ---
        $andereOU       = ""
        $statusGeloescht = ""
        $userGroups     = @()

        if ($ADObj) {
            $dn = $ADObj.DistinguishedName
            if ($dn -notmatch "OU=81" -and $dn -notmatch "OU=82") {
                $andereOU = $dn -replace '^CN=.*?,'
            }
            if ($ADObj.MemberOf) {
                $userGroups = @($ADObj.MemberOf | ForEach-Object {
                    ($_ -split ',')[0] -replace 'CN=',''
                })
            }
        } else {
            $statusGeloescht = "XXX"
        }

        # --- Alle bedingten Werte vorberechnen ---
        $isRequired  = if ($RequiredList.Contains($LID.ToUpper())) { "ZZZ" } else { "" }
        $isLafp      = ($LID -like "L110*") -or ($LID -like "L114*")
        $lafpStr     = if ($isLafp) { "LLL" } else { "" }
        $ouNum       = if ($isLafp) { "26"  } else { "[OU]" }

        $targetOrt = ""
        if ($CsvRow -and $CsvRow.Standort -and $CsvRow.Standort.ToString().Trim() -ne "") {
            $targetOrt = $CsvRow.Standort.ToString().Trim()
        } elseif ($ADObj -and $ADObj.l) {
            $targetOrt = $ADObj.l
        }

        $targetNachname = $targetOrt
        if ($isLafp -and $targetOrt -notmatch "^LAFP\s-\s") {
            $targetNachname = "LAFP - $targetOrt"
        }

        $buroPlatz = "[Platznummer]"
        if ($CsvRow -and $CsvRow.Raum -and $CsvRow.Raum.ToString().Trim() -ne "") {
            $buroPlatz = $CsvRow.Raum.ToString().Trim()
        }

        $fortbildung = ""
        if ($CsvRow -and $CsvRow.Fortbildungsbereich) {
            $fortbildung = $CsvRow.Fortbildungsbereich.ToString().Trim()
        }

        $anmerkung = ""
        if ($CsvRow -and $CsvRow.Anmerkungen) {
            $anmerkung = $CsvRow.Anmerkungen.ToString().Trim()
        }

        $aenderInfo = if ($fortbildung -and $anmerkung) {
            "$fortbildung - $anmerkung"
        } else {
            "$fortbildung$anmerkung"
        }

        $verfStr = GetVerfahren -Row $CsvRow

        $adVorname    = if ($ADObj) { "$($ADObj.GivenName)"               } else { "" }
        $adNachname   = if ($ADObj) { "$($ADObj.Surname)"                 } else { "" }
        $adDisplay    = if ($ADObj) { "$($ADObj.DisplayName)"             } else { "" }
        $adOrt        = if ($ADObj) { "$($ADObj.l)"                       } else { "" }
        $adBuero      = if ($ADObj) { "$($ADObj.physicalDeliveryOfficeName)" } else { "" }
        $adDez        = if ($ADObj) { "$($ADObj.department)"              } else { "" }
        $adDesc       = if ($ADObj) { "$($ADObj.Description)"             } else { "" }

        $origStandort = if ($CsvRow) { "$($CsvRow.Standort)" } else { "" }
        $origRaum     = if ($CsvRow) { "$($CsvRow.Raum)"     } else { "" }
        $origAbt      = if ($CsvRow) { "$($CsvRow.Abteilung)"} else { "" }

        $aendernDez   = if ($CsvRow -and $ADObj) {
            "$origAbt - $adDez"
        } else {
            $origAbt
        }

        $aendernDesc  = "$ouNum - $verfStr - $buroPlatz - [Verantwortlicher] - [TEL] | $aenderInfo"
        $displayName  = "$LID $targetNachname"

        # --- Hashtable OHNE inline-if (FIX-7) ---
        $Props = [ordered]@{
            "L-Kennung"                      = $LID
            "andere OU"                      = $andereOU
            "GELOESCHT"                      = $statusGeloescht
            "Benoetigt"                      = $isRequired
            "Geaendert"                      = ""
            "LAFP_LZPD_LKA"                  = $lafpStr
            "NICHT_KONFORM"                  = ""
            "AD_Vorname"                     = $adVorname
            "AENDERN_Vorname"                = $LID
            "AD_Nachname"                    = $adNachname
            "AENDERN_Nachname"               = $targetNachname
            "AD_DisplayName"                 = $adDisplay
            "AENDERN_DisplayName"            = $displayName
            "ORIGINAL_Standort"              = $origStandort
            "ORIGINAL_Raum_Schulungskreis"   = $origRaum
            "AD_Ort"                         = $adOrt
            "AENDERN_Ort"                    = $targetOrt
            "AD_Buero"                       = $adBuero
            "AENDERN_Buero"                  = $buroPlatz
            "ORIGINAL_Abteilung"             = $origAbt
            "AD_Dez"                         = $adDez
            "AENDERN_Dez"                    = $aendernDez
            "AD_Description"                 = $adDesc
            "AENDERN_Description"            = $aendernDesc
            "AENDERN_OU"                     = $ouNum
            "AENDERN_Info"                   = $aenderInfo
            "LOESCHEN"                       = ""
        }

        # --- Konformitaetspruefung ---
        if ($ADObj) {
            $codes = [System.Collections.Generic.List[string]]::new()
            if (IsDiff $adVorname  $LID)           { $codes.Add("VN")   }
            if (IsDiff $adNachname $targetNachname) { $codes.Add("NN")   }
            if (IsDiff $adDisplay  $displayName)    { $codes.Add("DN")   }
            if (IsDiff $adDesc     $aendernDesc)    { $codes.Add("DESC") }
            if (IsDiff $adOrt      $targetOrt)      { $codes.Add("ORT")  }
            if (IsDiff $adBuero    $buroPlatz)      { $codes.Add("GEB")  }
            if (IsDiff $adDez      $aendernDez)     { $codes.Add("DEZ")  }
            if (IsDiff $ADObj.info $aenderInfo)     { $codes.Add("INFO") }

            if ($codes.Count -gt 0) {
                $codeStr              = $codes -join ","
                $Props."Geaendert"    = $codeStr
                $Props."NICHT_KONFORM" = $codeStr
            }
        }

        # --- Gruppenmitgliedschaften ---
        foreach ($gn in $SortedGroups) {
            $Props["GRP_$gn"] = if ($userGroups -contains $gn) { "X" } else { "" }
        }

        return [PSCustomObject]$Props
    }
    # -----------------------------------------------------------------------

    # Jobs starten
    $Jobs = [System.Collections.Generic.List[PSObject]]::new()
    foreach ($LID in $ProcessList) {
        $psi = [powershell]::Create()
        [void]$psi.AddScript($ScriptBlock)
        [void]$psi.AddArgument($LID)
        [void]$psi.AddArgument($MasterCsvData[$LID])
        [void]$psi.AddArgument($ADCache[$LID])
        [void]$psi.AddArgument($RequiredList)
        [void]$psi.AddArgument($SortedGroups)
        $psi.RunspacePool = $Pool
        [void]$Jobs.Add([PSCustomObject]@{
            SAM         = $LID
            Instance    = $psi
            AsyncResult = $psi.BeginInvoke()
        })
    }

    # Jobs einsammeln
    $Results  = [System.Collections.Generic.List[PSObject]]::new()
    $Count    = 0
    $ErrCount = 0
    $Total    = $Jobs.Count

    while ($Jobs.Count -gt 0) {
        $Finished = @($Jobs | Where-Object { $_.AsyncResult.IsCompleted })
        $ToRemove = [System.Collections.Generic.List[PSObject]]::new()

        foreach ($Job in $Finished) {
            $Count++

            # Fehler-Stream ausgeben (OPT-1)
            foreach ($e in $Job.Instance.Streams.Error) {
                $ErrCount++
                $em = "[$( Get-Date -Format 'HH:mm:ss')] [ERROR] Runspace $($Job.SAM): $e"
                Write-Host $em -ForegroundColor Red
                $em | Out-File $LogFile -Append
            }

            # FIX-2: Robustes Unrolling - nur PSCustomObject-artige Objekte
            $rawOut = $Job.Instance.EndInvoke($Job.AsyncResult)
            foreach ($item in $rawOut) {
                if ($null -ne $item -and
                    $item.PSObject -ne $null -and
                    $item -isnot [System.Management.Automation.PSDataCollection[psobject]]) {
                    [void]$Results.Add($item)
                }
            }

            Write-Host "[$Count/$Total] OK: $($Job.SAM)" -ForegroundColor Gray
            Write-Progress -Activity "Verarbeite L-Kennungen" `
                           -Status "$Count / $Total" `
                           -PercentComplete ([math]::Round(($Count / $Total) * 100))

            $Job.Instance.Dispose()
            [void]$ToRemove.Add($Job)
        }

        # FIX-4: Erst nach Loop entfernen
        foreach ($j in $ToRemove) { [void]$Jobs.Remove($j) }
        if ($Jobs.Count -gt 0) { Start-Sleep -Milliseconds 50 }
    }

    Write-Progress -Activity "Verarbeite L-Kennungen" -Completed
    $Pool.Close(); $Pool.Dispose()

    return @{
        Results = $Results
        ErrCount = $ErrCount
    }
}

function Export-Results {
    param(
        [System.Collections.Generic.List[PSObject]]$Results,
        [array]$SortedGroups,
        [string]$Version,
        [string]$ScriptDir
    )
    # 5. EXPORT
    # FIX-6: Bounds-Check
    if ($Results.Count -eq 0) {
        Write-Log "FEHLER: Keine Ergebnisse generiert." "ERROR"; return
    }

    Write-Log "Erzeuge Ausgabedateien ($($Results.Count) Datensaetze)..." "INFO"

    # Tabelle 1: Gruppen-Uebersicht
    $Path1    = Join-Path $ScriptDir "L-Kennungen_Full_Analysis_v$Version.csv"
    $baseCols = @("L-Kennung","andere OU","GELOESCHT","Benoetigt","Geaendert","LOESCHEN",
                  "AD_Vorname","AD_Nachname")
    $grpCols  = $SortedGroups | ForEach-Object { "GRP_$_" }
    $Results | Select-Object ($baseCols + $grpCols) |
        Export-Csv -Path $Path1 -Delimiter ';' -NoTypeInformation -Encoding UTF8

    # Tabelle 2: Properties (alle Spalten ausser LOESCHEN, mit RFC-4180-Escaping)
    $Path2    = Join-Path $ScriptDir "L-Kennungen_Properties_v$Version.csv"
    $allCols  = $Results[0].psobject.Properties.Name | Where-Object { $_ -ne "LOESCHEN" }
    $lines    = [System.Collections.Generic.List[string]]::new()
    $lines.Add($allCols -join ';')
    foreach ($r in $Results) {
        $row = foreach ($col in $allCols) {
            $val = $r.$col
            $str = if ($null -ne $val) { $val.ToString() } else { "" }
            if ($str -match '[;\r\n"]') { '"' + $str.Replace('"','""') + '"' } else { $str }
        }
        $lines.Add($row -join ';')
    }
    $lines | Out-File -FilePath $Path2 -Encoding UTF8 -Force

    return @{
        Path1 = $Path1
        Path2 = $Path2
    }
}

function Invoke-Main {
    Show-Header
    Import-Module ActiveDirectory

    if (-not $CsvPath) { $CsvPath = Read-Host "Pfad zur Master-CSV eingeben" }
    $CsvPath = $CsvPath.Trim().Trim('"')
    if (-not (Test-Path $CsvPath)) {
        Write-Log "Master-CSV nicht gefunden: '$CsvPath'" "ERROR"; return
    }

    $RequiredList = Get-RequiredList -RequiredCsvPath $RequiredCsvPath
    $MasterCsvData = Get-MasterCsvData -CsvPath $CsvPath

    $adData = Get-ADData -SearchGlobal:$SearchGlobal -MasterCsvData $MasterCsvData -TestCount $TestCount
    $ADCache = $adData.ADCache
    $SortedGroups = $adData.SortedGroups
    $ProcessList = $adData.ProcessList

    $processResult = Process-Runspaces -MaxThreads $MaxThreads -ProcessList $ProcessList -MasterCsvData $MasterCsvData -ADCache $ADCache -RequiredList $RequiredList -SortedGroups $SortedGroups -LogFile $LogFile
    $Results = $processResult.Results
    $ErrCount = $processResult.ErrCount

    if ($Results.Count -gt 0) {
        $exportResult = Export-Results -Results $Results -SortedGroups $SortedGroups -Version $Version -ScriptDir $ScriptDir
        $Path1 = $exportResult.Path1
        $Path2 = $exportResult.Path2
    }

    $Stopwatch.Stop()

    Write-Host "`n====================================================" -ForegroundColor Cyan
    Write-Host "ERGEBNISSE:" -ForegroundColor Yellow
    Write-Host "VERSION : $Version"
    Write-Host "DATUM   : $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')"
    Write-Host "ZEILEN  : $($Results.Count)"
    Write-Host "FEHLER  : $ErrCount"
    Write-Host "DAUER   : $($Stopwatch.Elapsed.TotalSeconds.ToString('F2'))s"
    if ($Results.Count -gt 0) {
        Write-Host "TAB 1   : $Path1"
        Write-Host "TAB 2   : $Path2"
    }
    Write-Host "LOG     : $LogFile"
    Write-Host "====================================================`n" -ForegroundColor Cyan
}

Invoke-Main
