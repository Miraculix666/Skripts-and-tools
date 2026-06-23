# AD L-Kennung Management Toolset

Toolset zur Analyse, Bereinigung und Aktualisierung von L-Kennungen im Active Directory
der Polizei NRW (OUs 81/82).

---

## Dateien

| Datei | Version | Zweck |
|-------|---------|-------|
| `PS_LAccount_Manager.ps1` | v10.0 | Analysiert AD, baut Vergleichstabelle, erkennt Nicht-Konformitaeten |
| `PS_LAccount_Apply.ps1`   | v2.0  | Liest die Tabelle aus dem Manager und wendet Aenderungen auf AD an  |

---

## Workflow

```
1. PS_LAccount_Manager.ps1
   - Liest Master-CSV (Solldaten)
   - Liest AD (Istdaten) -- SearchScope Subtree auf OU=81 und OU=82
     -> alle Sub-OUs werden erfasst, auch unbekannte
   - Vergleicht, schreibt Codes in Spalte NICHT_KONFORM
   - Exportiert:
       L-Kennungen_Full_Analysis_v10.0.csv   <- Gruppen-Uebersicht
       L-Kennungen_Properties_v10.0.csv      <- Eingabe fuer Apply

2. Manuelle Pruefung der Properties-CSV (optional)
   - LOESCHEN-Spalte befuellen fuer Deaktivierungen
   - FILTER_<x>-Spalte fuer gezielte Selektion nutzen

3. PS_LAccount_Apply.ps1
   - Liest die Properties-CSV
   - Zeigt Aktionsplan + Bestaetigung
   - Setzt Properties, deaktiviert Konten, synchronisiert Gruppen
```

---

## PS_LAccount_Manager.ps1  (v10.0)

### Aenderungen gegenueber v9.6

| # | Aenderung |
|---|-----------|
| 1 | Parallel immer aktiv -- kein `-Parallel` Switch mehr |
| 2 | `GroupMode Single` ist Standard (alle Gruppen in einer Zelle) |
| 3 | Kein `GRP_`-Prefix auf Gruppenspaltennamen |
| 4 | Spaltennamen direkt kompatibel mit `PS_LAccount_Apply v2.0` |
| 5 | `SearchScope Subtree` auf OU=81 und OU=82 -- alle Sub-OUs immer erfasst |
| 6 | Neue Spalten `OU` (unmittelbarer Container) und `OU_Pfad` (vollst. DN) |
| 7 | `-GroupFilter`: erzeugt Markerspalte `FILTER_<Wert>` fuer Selektion |
| 8 | `-OUFilter`: beschleunigt Discovery auf Teilmenge von OUs |
| 9 | `-NameFilter`: beschleunigt Verarbeitung auf bestimmte SAM-Namensmuster |
| 10 | `-LafpOnly`: verarbeitet nur LAFP-Kennungen (L110* / L114*) |

### Parameter

| Parameter | Typ | Pflicht | Beschreibung |
|-----------|-----|---------|--------------|
| `-CsvPath` | String | Ja | Pfad zur Master-CSV (Solldaten, Semikolon-getrennt) |
| `-RequiredCsvPath` | String | Nein | Pfad zur Bedarfsliste (markiert Kennungen als ZZZ) |
| `-MaxThreads` | Int | Nein | Anzahl paralleler Threads (Default: 8) |
| `-GroupMode` | String | Nein | `Single` = alle Gruppen in einer Zelle (Standard) / `Columns` = eine Spalte pro Gruppe |
| `-GroupFilter` | String | Nein | Teilstring (case-insensitiv): Kennungen deren Gruppen passen erhalten X in Spalte `FILTER_<Wert>`. Dient der Selektion in Excel / Apply. |
| `-OUFilter` | String | Nein | Nur User aus OUs deren DN diesen Teilstring enthaelt. Wird waehrend Discovery angewendet -- beschleunigt Scan. |
| `-NameFilter` | String | Nein | SAMAccountName-Wildcard. Wird vor Verarbeitung angewendet -- beschleunigt Lauf. |
| `-LafpOnly` | Switch | Nein | Nur LAFP-Kennungen (L110* und L114*). Kombinierbar mit `-NameFilter`. |
| `-TestCount` | Int | Nein | Nur erste N Eintraege verarbeiten (Testmodus) |
| `-DebugMode` | Switch | Nein | Ausfuehrliches Logging pro Eintrag |

### Beispiele

```powershell
# Standardlauf (parallel, GroupMode Single):
.\PS_LAccount_Manager.ps1 -CsvPath .\master.csv

# Nur OU 81, Gruppen mit "Schulung" markieren:
.\PS_LAccount_Manager.ps1 -CsvPath .\master.csv -OUFilter "OU=81" -GroupFilter "Schulung"

# Nur LAFP-Konten, Gruppen als Spalten:
.\PS_LAccount_Manager.ps1 -CsvPath .\master.csv -LafpOnly -GroupMode Columns

# Test-Gruppen markieren (Spalte FILTER_test in beiden CSVs):
.\PS_LAccount_Manager.ps1 -CsvPath .\master.csv -GroupFilter "test"

# Nur Kennungen die "1234" im Namen enthalten:
.\PS_LAccount_Manager.ps1 -CsvPath .\master.csv -NameFilter "*1234*"

# Mit Bedarfsliste, 16 Threads, Testlauf:
.\PS_LAccount_Manager.ps1 -CsvPath .\master.csv -RequiredCsvPath .\bedarf.csv -MaxThreads 16 -TestCount 50 -DebugMode
```

### Ausgabe-Spalten (Properties-CSV / Eingabe fuer Apply)

| Spaltengruppe | Spalten |
|---------------|---------|
| Identifikation | `L-Kennung`, `OU`, `OU_Pfad`, `andere_OU`, `GELOESCHT`, `Benoetigt`, `LAFP_LZPD_LKA` |
| Konformitaet | `NICHT_KONFORM`, `Geaendert` |
| AD Istwerte | `AD_Vorname`, `AD_Nachname`, `AD_DisplayName`, `AD_Ort`, `AD_Buero`, `AD_Dez`, `AD_Description` |
| Zielwerte (Apply) | `AENDERN_Vorname`, `AENDERN_Nachname`, `AENDERN_DisplayName`, `AENDERN_Ort`, `AENDERN_Buero`, `AENDERN_Dez`, `AENDERN_Description`, `AENDERN_Info`, `AENDERN_OU` |
| Steuerung | `LOESCHEN` (manuell befuellen fuer Deaktivierung) |
| Selektion | `FILTER_<Wert>` (nur wenn `-GroupFilter` gesetzt, X = mind. eine passende Gruppe) |
| Gruppen | `Gruppen` (Single) oder `Gruppenname` pro Spalte (Columns) |

### NICHT_KONFORM Codes

| Code | AD-Attribut | Quellspalte |
|------|-------------|-------------|
| VN   | GivenName   | AENDERN_Vorname |
| NN   | Surname     | AENDERN_Nachname |
| DN   | DisplayName | AENDERN_DisplayName |
| DESC | Description | AENDERN_Description |
| ORT  | City / l    | AENDERN_Ort |
| GEB  | Office      | AENDERN_Buero |
| DEZ  | Department  | AENDERN_Dez |
| INFO | info        | AENDERN_Info |

### GroupFilter-Logik (Manager)

Der `-GroupFilter` Parameter erzeugt eine zusaetzliche Spalte `FILTER_<Wert>` in beiden Ausgabe-CSVs.
Ein `X` wird gesetzt wenn mindestens eine Gruppe des Benutzers den Teilstring enthaelt (case-insensitiv).

```
Benutzer       Gruppen                          GroupFilter="test"   FILTER_test
---------------------------------------------------------------------------
L1001000       OU81-Testgruppe, OU81-Verwaltung  passt               X
L1001001       OU81-Schulung, OU81-Verwaltung    kein Treffer
L1001002       TestAdmins, OU82-AB-Test          2 Treffer           X
```

Die Spalte dient als Selektionshilfe -- in Excel filtern oder im Apply als Grundlage nutzen.

---

## PS_LAccount_Apply.ps1  (v2.0)

### Parameter

| Parameter | Typ | Pflicht | Beschreibung |
|-----------|-----|---------|--------------|
| `-CsvPath` | String | Ja | Pfad zur Properties-CSV (Ausgabe des Managers) |
| `-ColFilter` | String | Nein | Komma-/Semikolonliste von Spaltennamen, die verarbeitet werden sollen. Leer = alle NICHT_KONFORM-Spalten. Sonderfall: `LOESCHEN` aktiviert nur Deaktivierungen. |
| `-GroupSync` | Switch | Nein | Gruppen-Sync ausfuehren (Default: aus) |
| `-GroupFilter` | String | Nein | Substring-Filter (case-insensitiv) fuer Gruppennamen. Nur Gruppen deren Name passt werden verarbeitet. |
| `-WhatIf` | Switch | Nein | Trockenlauf: zeigt alles, aendert nichts |
| `-DebugMode` | Switch | Nein | Ausfuehrliches Logging |

### Beispiele

```powershell
# Trockenlauf -- alles pruefen ohne zu aendern:
.\PS_LAccount_Apply.ps1 -CsvPath .\Properties.csv -WhatIf

# Alle nicht-konformen Properties anwenden:
.\PS_LAccount_Apply.ps1 -CsvPath .\Properties.csv

# Nur Vor-/Nachname + DisplayName korrigieren:
.\PS_LAccount_Apply.ps1 -CsvPath .\Properties.csv -ColFilter "AENDERN_Vorname,AENDERN_Nachname,AENDERN_DisplayName"

# Nur Deaktivierungen:
.\PS_LAccount_Apply.ps1 -CsvPath .\Properties.csv -ColFilter "LOESCHEN"

# Gruppen-Sync, nur Gruppen mit "Schulung" im Namen:
.\PS_LAccount_Apply.ps1 -CsvPath .\Properties.csv -GroupSync -GroupFilter "Schulung"

# Test-Gruppen synchronisieren:
.\PS_LAccount_Apply.ps1 -CsvPath .\Properties.csv -GroupSync -GroupFilter "test"

# Alles inklusive Gruppen:
.\PS_LAccount_Apply.ps1 -CsvPath .\Properties.csv -GroupSync
```

### GroupFilter-Logik (Apply)

Im Apply steuert `-GroupFilter` welche Gruppen bei der Synchronisation beruecksichtigt werden.
Nur Gruppen deren Name den Teilstring enthaelt werden hinzugefuegt oder entfernt.

```
Gruppenname            GroupFilter="test"    Aktion
-------------------------------------------------------
OU81-AB-Testgruppe     passt (*test*)        Add / Remove
OU81-AB-Schulung       kein Treffer          ignoriert
OU82-TestAdmins        passt (*test*)        Add / Remove
```

### Aktionstypen

| Typ | Ausloeser | AD-Aktion |
|-----|-----------|-----------|
| UPDATE | NICHT_KONFORM enthaelt Codes | Set-ADUser |
| DEAKTIV | LOESCHEN-Spalte befuellt | Disable-ADAccount |
| GRP-SYNC | `-GroupSync` + Gruppenveraenderungen | Add/Remove-ADGroupMember |
| SKIP | Keine Aktion / gefiltert | -- |

---

## Gruppen-Spaltenformat

| GroupMode im Manager | Spaltenformat in CSV | Auswertung im Apply |
|----------------------|----------------------|---------------------|
| `Single` (Standard) | Spalte `Gruppen` (semikolon-separiert) | enthaltene Gruppen = hinzufuegen |
| `Columns` | Spaltenname = Gruppenname (kein GRP_-Prefix) | X = hinzufuegen, leer = entfernen |

Hinweis: Bei `GroupMode Single` werden im Apply nur Gruppen hinzugefuegt, nicht entfernt.
Fuer vollstaendigen Gruppen-Sync (Add + Remove) wird `GroupMode Columns` empfohlen.

---

## Ausgabedateien

### Manager
| Datei | Inhalt |
|-------|--------|
| `L-Kennungen_Full_Analysis_v10.0.csv` | Gruppen-Uebersicht (kompakt, FILTER_-Spalte wenn gesetzt) |
| `L-Kennungen_Properties_v10.0.csv` | Alle Properties -- Eingabedatei fuer Apply |
| `AD_Sync_<Timestamp>.log` | Lauf-Protokoll |

### Apply
| Datei | Inhalt |
|-------|--------|
| `Apply_Result_<Timestamp>.csv` | Verarbeitungsergebnis je L-Kennung |
| `Apply_Log_<Timestamp>.log` | Lauf-Protokoll |

#### Spalten Apply_Result_*.csv

| Spalte | Beschreibung |
|--------|--------------|
| LID | L-Kennung |
| Aktion | UPDATE / DEAKTIV / GRP-SYNC |
| Codes | Verarbeitete NICHT_KONFORM-Codes |
| Loeschen | Inhalt der LOESCHEN-Spalte |
| Status | OK-UPDATE / OK-DEAKTIV / OK-GRP / FEHLER / SKIP-* |
| Details | Geaenderte Attribute und Werte |
| Timestamp | Zeitstempel |

---

## Voraussetzungen

- Windows PowerShell 5.1 oder PowerShell 7+
- RSAT: Active Directory Modul (`Import-Module ActiveDirectory`)
- Schreibrechte auf den betreffenden OUs (fuer Apply)

## Bekannte Einschraenkungen

- Hard-Delete nicht implementiert: `LOESCHEN` deaktiviert nur das Konto
- OU-Verschiebung (`AENDERN_OU`) noch nicht automatisiert
- `GroupMode Single`: im Apply nur Hinzufuegen, kein Entfernen moeglich
