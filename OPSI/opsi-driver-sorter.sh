#!/bin/bash

# ==============================================================================
# OPSI All-in-One Treiber-Skript
#
# Zweck:      Sortiert Windows-Treiber robust in die OPSI-Struktur für ein
#             beliebiges Netboot-Produkt. Nutzt eine Cache-DB, liest optional
#             Audits, vergleicht Versionen und verwendet absolute Pfade.
#
# Autor:      PS-Coding
# Version:    28.0 - Stabile Client-Abfrage & finale Korrekturen
# Datum:      2025-09-22
#
# ANWENDUNG:
# Das Skript kann nun eine Produkt-ID als Parameter annehmen.
#
# Standard (win10-x64):
#   sudo ./opsi-driver-sorter.sh
#
# Für ein anderes Produkt (z.B. win11-x64):
#   sudo ./opsi-driver-sorter.sh -p win11-x64
#
# Hilfe anzeigen:
#   ./opsi-driver-sorter.sh -h
# ==============================================================================

# --- KONFIGURATION ---
# Standard-Produkt-ID, falls keine über Parameter übergeben wird.
DEFAULT_PRODUCT_ID="win10-x64"

# true: Nur Simulation, keine Änderungen am System.
# false: Skript führt alle Aktionen aus.
DRY_RUN=false

# true: Zeigt jeden Befehl an, der ausgeführt wird (für detailliertes Debugging).
# false: Normale Ausgabe.
DEBUG=true

# true: Erstellt nach dem Kopieren automatisch die .opsi-Treiberpakete.
# false: Überspringt die Paketerstellung.
CREATE_DRIVER_PACKAGES=true

# Steuert das Verhalten zum Löschen der Quell-Treiber nach dem Kopieren.
# Mögliche Werte:
#   "per_file":   (Standard) Löscht jedes Quell-Verzeichnis sofort nach dem erfolgreichen Kopieren.
#   "on_success": Löscht ALLE erfolgreich kopierten Quell-Verzeichnisse am Ende des Skripts.
#   "never":      Behält alle Quell-Verzeichnisse bei.
DELETE_SOURCE_AFTER_COPY="per_file"
# --- ENDE DER KONFIGURATION ---


# --- Skript-Logik (bitte nicht ändern) ---

# Stellt die UTF-8 Kodierung für das gesamte Skript sicher, um Fehler bei der Log-Ausgabe zu vermeiden.
export LANG=C.UTF-8

# Farben für die Ausgabe
GREEN='\033[1;32m'; YELLOW='\033[1;33m'; BLUE='\033[1;34m'; RED='\033[1;31m'; NC='\033[0m'

# Zentrale Logging-Funktion (wird vor dem Logging-Setup definiert, um für Parameter-Parsing verfügbar zu sein)
log_message() {
    # Wenn LOG_FILE noch nicht gesetzt ist, nur in die Konsole ausgeben
    if [ -z "$LOG_FILE" ]; then
        echo -e "${2}"
        return
    fi
    local level=$1; local message=$2; local color=$NC
    local plain_message="[${level}] ${message}"
    case "$level" in
        INFO) color=$BLUE ;; OK) color=$GREEN ;; WARN) color=$YELLOW ;; ERROR) color=$RED ;; COPY) color=$GREEN ;;
    esac
    # Schreibt die Nachricht in die Konsole und in die Log-Datei
    echo -e "${color}${plain_message}${NC}" | tee -a "$LOG_FILE"
}

# --- HILFEFUNKTION UND PARAMETER-PARSING ---
show_help() {
    echo "OPSI All-in-One Treiber-Skript V28.0"
    echo ""
    echo "Anwendung: $0 [-p PRODUKT_ID]"
    echo ""
    echo "Optionen:"
    echo "  -p, --product   Die OPSI-Produkt-ID des zu bearbeitenden Netboot-Produkts."
    echo "                  Standardwert, falls nicht angegeben: '${DEFAULT_PRODUCT_ID}'"
    echo "  -h, --help        Diese Hilfe anzeigen."
    echo ""
}

PRODUCT_ID="$DEFAULT_PRODUCT_ID"

# Argumente verarbeiten
while [[ "$#" -gt 0 ]]; do
    case $1 in
        -p|--product)
            if [[ -n "$2" && ! "$2" =~ ^- ]]; then
                PRODUCT_ID="$2"
                shift 2
            else
                log_message "ERROR" "Fehler: Option '-p' erfordert ein Argument." >&2
                exit 1
            fi
            ;;
        -h|--help)
            show_help
            exit 0
            ;;
        *) # Unbekannte Option
            log_message "ERROR" "Unbekannte Option: $1" >&2
            show_help
            exit 1
            ;;
    esac
done


# --- DYNAMISCHE PFADE & VARIABLEN ---
OPSI_DEPOT_PATH="/var/lib/opsi/depot"
PRODUCT_DEPOT_PATH="${OPSI_DEPOT_PATH}/${PRODUCT_ID}"
TARGET_DRIVER_DIR="${PRODUCT_DEPOT_PATH}/drivers/drivers"
UNSORTED_DRIVER_DIR="${PRODUCT_DEPOT_PATH}/drivers_source"
CACHE_DB_FILE="${PRODUCT_DEPOT_PATH}/drivers/driver.db" # Cache-DB liegt eine Ebene höher
AUDIT_CACHE_FILE="${PRODUCT_DEPOT_PATH}/drivers/audit_hw_ids.cache" # Cache für Audit-Ergebnisse

# Definition der Befehlspfade
OPSI_ADMIN_CMD="/usr/bin/opsi-admin"
OPSI_SET_RIGHTS_CMD="/usr/bin/opsi-set-rights"
CREATE_DRIVER_LINKS_CMD="${PRODUCT_DEPOT_PATH}/create_driver_links.py"
OPSI_PKG_MGR_CMD="/usr/bin/opsi-package-manager"
OPSI_SETUP_CMD="/usr/bin/opsi-setup"

# Assoziative Arrays und Listen für die Logik
declare -A LATEST_DRIVERS
declare -A NEEDED_HW_IDS
declare -a SUCCESSFULLY_COPIED_SOURCES=()
INTEGRATE_ONLY_NEEDED_DRIVERS=false

# --- Logging-Setup (Jetzt, da PRODUCT_ID final ist) ---
LOG_DIR="/var/log/opsi"
LOG_FILE="${LOG_DIR}/driver_sort_${PRODUCT_ID}_$(date +%Y-%m-%d_%H-%M-%S).log"
mkdir -p "$LOG_DIR"
touch "$LOG_FILE"


if [[ $DEBUG = true ]]; then
    set -x
fi

# --- Skript-Start ---
log_message "INFO" "Starte OPSI Treiber-Sortierungsskript V28.0"
log_message "INFO" "Verarbeite Produkt-ID: ${PRODUCT_ID}"
log_message "INFO" "Log-Datei: ${LOG_FILE}"


# Funktion für optische Trennlinien in der Ausgabe
print_header() {
    log_message "INFO" "=================================================================="
    log_message "OK"   "$1"
    log_message "INFO" "=================================================================="
}

# 1. Start und grundlegende Überprüfungen
if [[ $DRY_RUN = true ]]; then
    print_header "!!! ACHTUNG: TROCKENLAUF (DRY RUN) IST AKTIV !!!"
    log_message "WARN" "Es werden keine Änderungen vorgenommen."
fi

if [[ $EUID -ne 0 ]]; then
   log_message "ERROR" "Dieses Skript muss mit root-Rechten ausgeführt werden."; exit 1
fi
if [ ! -d "$PRODUCT_DEPOT_PATH" ]; then
    log_message "ERROR" "Produktverzeichnis für '${PRODUCT_ID}' nicht gefunden: ${PRODUCT_DEPOT_PATH}"; exit 1
fi
if [ ! -d "$UNSORTED_DRIVER_DIR" ]; then
    log_message "ERROR" "Treiber-Quellverzeichnis nicht gefunden: ${UNSORTED_DRIVER_DIR}"; exit 1
fi
# Stellt sicher, dass das Ziel- und Cache-Verzeichnis existiert
mkdir -p "$TARGET_DRIVER_DIR"
mkdir -p "$(dirname "$CACHE_DB_FILE")"
touch "$CACHE_DB_FILE"

# HILFSFUNKTION: Konvertiert ein Treiberdatum und eine Version in eine vergleichbare Zahl
normalize_version() {
    local date_str version_str y m d a b c d
    IFS=',' read -r date_str version_str <<< "$1"
    IFS='/' read -r m d y <<< "$date_str"
    IFS='.' read -r a b c d <<< "$version_str"
    printf "%04d%02d%02d%04d%04d%06d%05d" "${y:-0}" "${m:-0}" "${d:-0}" "${a:-0}" "${b:-0}" "${c:-0}" "${d:-0}"
}

# HILFSFUNKTION: Liest den Inhalt einer INF-Datei mit Fallback für die Kodierung
read_inf_content() {
    local file_path="$1"
    # Versucht die Konvertierung, wenn diese fehlschlägt, wird die Datei als Standardtext gelesen.
    iconv -f UTF-16LE -t UTF-8 "$file_path" 2>/dev/null || cat "$file_path" 2>/dev/null
}

# PHASE 1: AUDIT-DATEN SAMMELN (OPTIONAL)
print_header "Phase 1/5: Auswahl der Integrationsmethode"
read -p "Sollen nur Treiber integriert werden, die laut OPSI-Audit benötigt werden? (j/N): " -n 1 -r; echo
if [[ $REPLY =~ ^[Jj]$ ]]; then
    INTEGRATE_ONLY_NEEDED_DRIVERS=true
    log_message "OK" "Selektive Integration basierend auf Audit-Daten aktiviert."

    PERFORM_NEW_AUDIT=false
    if [ ! -f "$AUDIT_CACHE_FILE" ]; then
        PERFORM_NEW_AUDIT=true
        log_message "WARN" "Keine Audit-Cache-Datei gefunden. Ein neues Audit ist erforderlich."
    elif [ -n "$(find "$AUDIT_CACHE_FILE" -mtime +365)" ]; then
        PERFORM_NEW_AUDIT=true
        log_message "WARN" "Die Audit-Cache-Datei ist älter als ein Jahr. Ein neues Audit wird empfohlen."
    else
        read -p "Eine aktuelle Audit-Cache-Datei wurde gefunden. Möchten Sie sie verwenden oder ein neues Audit erzwingen? (v)erwenden / (n)eu: " -n 1 -r; echo
        if [[ $REPLY =~ ^[Nn]$ ]]; then
            PERFORM_NEW_AUDIT=true
        fi
    fi

    if [ "$PERFORM_NEW_AUDIT" = true ]; then
        if [ ! -x "$OPSI_ADMIN_CMD" ]; then
            log_message "ERROR" "Befehl '${OPSI_ADMIN_CMD}' nicht gefunden. Audit-Analyse nicht möglich."
            exit 1
        fi
        
        log_message "INFO" "Ermittle Client-Liste für neues Audit..."
        client_ids=$($OPSI_ADMIN_CMD method getClientIds_list | tr -d '[],"')
        client_count=$(echo "$client_ids" | wc -w)

        if [ "$client_count" -eq 0 ]; then
            log_message "ERROR" "Keine OPSI-Clients in der Datenbank gefunden."
            exit 1
        fi
        
        read -p "Ein neues Audit wird für ${client_count} Clients durchgeführt (nur Clients mit vorhandenen HW-Daten liefern Ergebnisse). Fortfahren? (j/N): " -n 1 -r; echo
        if [[ $REPLY =~ ^[Jj]$ ]]; then
            log_message "INFO" "Starte neues Audit. Sammle Hardware-Daten (dies kann dauern)..."
            processed_clients=0
            TEMP_AUDIT_CACHE=$(mktemp)
            
            for client_id in $client_ids; do
                ((processed_clients++))
                log_message "INFO" "Verarbeite Client ${processed_clients}/${client_count}: ${client_id}"
                hw_data=$($OPSI_ADMIN_CMD method getHostHardware_hash "$client_id")
                
                if [ -z "$hw_data" ]; then
                    log_message "WARN" "Keine Hardware-Daten für Client ${client_id} erhalten. Wird übersprungen."
                    continue
                fi

                extracted_ids=$(echo "$hw_data" | tr -d '\0' | grep -o -i -E "(PCI|USB)\\\\[^]]*")

                echo "$extracted_ids" | while read -r line; do
                    shopt -s nocasematch
                    if [[ "$line" =~ (VEN_([0-9A-F]{4})) ]] && [[ "$line" =~ (DEV_([0-9A-F]{4})) ]]; then
                        echo "PCI\\${BASH_REMATCH[1]}&${BASH_REMATCH[3]}" >> "$TEMP_AUDIT_CACHE"
                    elif [[ "$line" =~ (VID_([0-9A-F]{4})) ]] && [[ "$line" =~ (PID_([0-9A-F]{4})) ]]; then
                        echo "USB\\${BASH_REMATCH[1]}&${BASH_REMATCH[3]}" >> "$TEMP_AUDIT_CACHE"
                    fi
                    shopt -u nocasematch
                done
            done
            
            sort -u "$TEMP_AUDIT_CACHE" > "$AUDIT_CACHE_FILE"
            rm "$TEMP_AUDIT_CACHE"
            log_message "OK" "Neues Audit abgeschlossen und in Cache-Datei gespeichert."
        else
            log_message "WARN" "Neues Audit vom Benutzer abgebrochen."
            if [ ! -f "$AUDIT_CACHE_FILE" ]; then
                log_message "ERROR" "Keine existierende Cache-Datei gefunden. Audit-basierte Integration nicht möglich. Breche ab."
                exit 1
            fi
            log_message "INFO" "Fahre mit der alten, existierenden Cache-Datei fort."
        fi
    else
        log_message "OK" "Verwende existierende Audit-Cache-Datei."
    fi
    
    # Lese benötigte HW-IDs aus der (neuen oder alten) Cache-Datei
    if [ ! -s "$AUDIT_CACHE_FILE" ]; then # -s prüft, ob die Datei größer als 0 ist
        log_message "ERROR" "Die Audit-Cache-Datei ist leer. Es wurden keine Hardware-IDs gefunden. Bitte führen Sie das Skript erneut aus und erzwingen Sie ein neues Audit."
        exit 1
    fi
    log_message "INFO" "Lese benötigte Hardware-IDs aus dem Cache..."
    while read -r line; do
        NEEDED_HW_IDS["${line^^}"]=1
    done < "$AUDIT_CACHE_FILE"
    log_message "OK" "${#NEEDED_HW_IDS[@]} eindeutige, benötigte Hardware-IDs aus dem Cache geladen."
else
    log_message "INFO" "Standard-Integration: Alle gefundenen Treiber werden verarbeitet."
fi

# PHASE 2: CACHE-ABGLEICH
print_header "Phase 2/5: Gleiche Treiber-DB mit dem Dateisystem ab..."
TEMP_DB=$(mktemp)
entry_count=0; pruned_count=0
if [ -s "$CACHE_DB_FILE" ]; then
    while IFS='|' read -r inf_path checksum version hwids || [[ -n "$inf_path" ]]; do
        ((entry_count++))
        if [ -f "$inf_path" ]; then
            echo "$inf_path|$checksum|$version|$hwids" >> "$TEMP_DB"
        else
            log_message "WARN" "Entferne veralteten DB-Eintrag für nicht mehr existierende Datei: $inf_path"
            ((pruned_count++))
        fi
    done < "$CACHE_DB_FILE"
    mv "$TEMP_DB" "$CACHE_DB_FILE"
fi
log_message "OK" "Treiber-DB-Abgleich abgeschlossen. ${pruned_count} von ${entry_count} Einträgen entfernt."

# PHASE 3: ANALYSE NEUER TREIBER
print_header "Phase 3/5: Analysiere neue/geänderte Treiber..."
TOTAL_INF_FILES=$(find "${UNSORTED_DRIVER_DIR}" -type f -iname "*.inf" 2>/dev/null | wc -l)
log_message "OK" "${TOTAL_INF_FILES} .inf-Dateien zur Analyse gefunden."

COUNT_PROCESSED_INF=0
find "${UNSORTED_DRIVER_DIR}" -type f -iname "*.inf" -print0 2>/dev/null | while IFS= read -r -d '' inf_file; do
    ((COUNT_PROCESSED_INF++))
    log_message "INFO" "------------------------------------------------------------------"
    log_message "INFO" "Prüfe Datei (${COUNT_PROCESSED_INF}/${TOTAL_INF_FILES}): ${inf_file}"
    current_checksum=$(md5sum "$inf_file" | cut -d' ' -f1)
    cached_entry=$(grep -F "|$current_checksum|" "$CACHE_DB_FILE" || true)

    if [ -n "$cached_entry" ]; then
        log_message "INFO" "Treiber ist unverändert (Cache-Treffer via Checksum). Überspringe Analyse."
        continue
    fi
    sed -i -e "\|^${inf_file}|d" "$CACHE_DB_FILE"
    
    log_message "INFO" "Neue/geänderte Datei. Führe Analyse durch..."
    
    driver_ver_line=$(read_inf_content "$inf_file" | grep -E -i -m 1 "DriverVer")
    
    if [ -z "$driver_ver_line" ]; then
        log_message "WARN" "Keine 'DriverVer'-Zeile gefunden in: ${inf_file}. Überspringe."
        continue
    fi

    current_ver_str=$(echo "$driver_ver_line" | sed -n 's/.*DriverVer\s*=\s*//p' | tr -d '[:space:]"')
    current_ver_norm=$(normalize_version "$current_ver_str")
    log_message "INFO" "Version gefunden: ${current_ver_str}"

    hw_id_pattern='PCI\\VEN_[0-9A-F]{4}&DEV_[0-9A-F]{4}|USB\\VID_[0-9A-F]{4}&PID_[0-9A-F]{4}'
    ids=$(read_inf_content "$inf_file" | grep -E -o -i "$hw_id_pattern" | tr -d '\r' | sort -u)

    if [ -z "$ids" ]; then
        inf_class_line=$(read_inf_content "$inf_file" | grep -E -i -m 1 "Class=")
        class_name=$(echo "$inf_class_line" | sed 's/.*Class=\s*//' | tr -d '"')
        if [[ "$class_name" == "System" || "$class_name" == "SoftwareComponent" || "$class_name" == "Monitor" || "$class_name" == "Firmware" ]]; then
            log_message "INFO" "Datei als System-Komponente identifiziert (Klasse: $class_name). Wird separat behandelt."
            ids="SYS_COMPONENT"
        else
            log_message "WARN" "Keine unterstützten HW-IDs und keine bekannte System-Klasse gefunden. Überspringe."
            continue
        fi
    fi

    ids_comma_separated=$(echo "$ids" | tr '\n' ',' | sed 's/,$//')
    echo "${inf_file}|${current_checksum}|${current_ver_norm}|${ids_comma_separated}" >> "$CACHE_DB_FILE"
    log_message "OK" "Analyse erfolgreich. DB-Eintrag hinzugefügt/aktualisiert."
done

# PHASE 4: PLANUNG DER KOPIERAKTIONEN
print_header "Phase 4/5: Ermittle neueste & benötigte Treiber..."
while IFS='|' read -r inf_path checksum version hwids || [[ -n "$inf_path" ]]; do
    IFS=',' read -ra hwid_array <<< "$hwids"
    for id_line in "${hwid_array[@]}"; do
        id_line_upper=${id_line^^} 
        if [ "$INTEGRATE_ONLY_NEEDED_DRIVERS" = true ] && [ -z "${NEEDED_HW_IDS[$id_line_upper]}" ] && [ "$id_line_upper" != "SYS_COMPONENT" ]; then
            continue
        fi
        if [ -z "${LATEST_DRIVERS[$id_line_upper]}" ]; then
            LATEST_DRIVERS[$id_line_upper]="${version};${inf_path}"
        else
            existing_ver_norm=$(echo "${LATEST_DRIVERS[$id_line_upper]}" | cut -d';' -f1)
            if (( version > existing_ver_norm )); then
                LATEST_DRIVERS[$id_line_upper]="${version};${inf_path}"
            fi
        fi
    done
done < "$CACHE_DB_FILE"
log_message "OK" "Planung abgeschlossen. ${#LATEST_DRIVERS[@]} Treiber-Aktionen sind vorgemerkt."

# PHASE 5: AUSFÜHRUNG DER KOPIERAKTIONEN
print_header "Phase 5/5: Kopiere neueste & benötigte Treiber"
COUNT_SUCCESS=0
declare -A COPIED_DIRS # Verhindert doppeltes Kopieren
for id_line in "${!LATEST_DRIVERS[@]}"; do
    inf_file=$(echo "${LATEST_DRIVERS[$id_line]}" | cut -d';' -f2-)
    current_source_dir=$(dirname "${inf_file}")

    if [ -n "${COPIED_DIRS[${current_source_dir}]}" ]; then
        continue
    fi
    
    if [[ $id_line == "SYS_COMPONENT" ]]; then
        TYPE="sys"
        VENDOR="components"
        DEVICE=$(basename "$current_source_dir" | tr -cd '[:alnum:]._-' | cut -c 1-20)
    elif [[ $id_line =~ ^PCI\\VEN_([0-9A-F]{4})\&DEV_([0-9A-F]{4})$ ]]; then
        TYPE="pciids"; VENDOR=${BASH_REMATCH[1]}; DEVICE=${BASH_REMATCH[2]}
    elif [[ $id_line =~ ^USB\\VID_([0-9A-F]{4})\&PID_([0-9A-F]{4})$ ]]; then
        TYPE="usbids"; VENDOR=${BASH_REMATCH[1]}; DEVICE=${BASH_REMATCH[2]}
    else
        continue
    fi

    local short_int_mode="man"
    if [ "$INTEGRATE_ONLY_NEEDED_DRIVERS" = true ]; then
        short_int_mode="aud"
    fi

    local short_sort_type="add"
    if [[ "${inf_file}" == *"/preferred/"* ]]; then short_sort_type="pref"; fi
    if [[ "${inf_file}" == *"/excluded/"* ]]; then short_sort_type="excl"; fi

    local short_type="pci"
    if [[ $TYPE == "usbids" ]]; then short_type="usb"; fi
    if [[ $TYPE == "sys" ]]; then short_type="sys"; fi

    final_target_path="${TARGET_DRIVER_DIR}/${short_sort_type}/${short_int_mode}/${short_type}/${VENDOR^^}/${DEVICE^^}"

    if [ "$DRY_RUN" = true ]; then
        log_message "INFO" "[DRY RUN] Quelle: '${current_source_dir}'"
        log_message "INFO" "[DRY RUN] Ziel für ${id_line} wäre: '${final_target_path}'"
        log_message "WARN" "[DRY RUN] Lösch-Aktion für Quelle würde übersprungen."
    else
        log_message "INFO" "Erstelle Verzeichnis (falls nötig): ${final_target_path}"
        mkdir -p "$final_target_path"
        
        log_message "COPY" "[KOPIERE] '${current_source_dir}/' -> '${final_target_path}/'"
        rsync -a --no-perms --no-owner --no-group --delete "${current_source_dir}/" "${final_target_path}/"
        
        if [ $? -eq 0 ]; then
            COPIED_DIRS[${current_source_dir}]=1
            SUCCESSFULLY_COPIED_SOURCES+=("${current_source_dir}")
            ((COUNT_SUCCESS++))
            
            if [ "$DELETE_SOURCE_AFTER_COPY" = "per_file" ]; then
                log_message "INFO" "[LÖSCHE QUELLE] '${current_source_dir}/'"
                rm -rf "${current_source_dir}"
            fi
        else
            log_message "ERROR" "rsync-Fehler beim Kopieren von '${current_source_dir}'. Quelle wird NICHT gelöscht."
        fi
    fi
done

# --- FOLGEAKTIONEN NACH DEM KOPIEREN ---
print_header "Schritt 6/9: Dateirechte werden korrigiert"
if [ "$DRY_RUN" = true ]; then
    log_message "WARN" "[DRY RUN] Befehl würde ausgeführt: ${OPSI_SET_RIGHTS_CMD} '${PRODUCT_DEPOT_PATH}'"
else
    if [ "$COUNT_SUCCESS" -gt 0 ]; then
        log_message "INFO" "Führe '${OPSI_SET_RIGHTS_CMD}' für '${PRODUCT_ID}' aus..."
        "$OPSI_SET_RIGHTS_CMD" "${PRODUCT_DEPOT_PATH}"
        log_message "OK" "Dateirechte wurden erfolgreich gesetzt."
    else
        log_message "WARN" "Keine Treiber kopiert, Korrektur der Rechte übersprungen."
    fi
fi

if [ "$CREATE_DRIVER_PACKAGES" = true ]; then
    print_header "Schritt 7/9: OPSI-Treiberpakete werden erstellt"
    if ! [ -f "$CREATE_DRIVER_LINKS_CMD" ]; then
        log_message "ERROR" "Skript '${CREATE_DRIVER_LINKS_CMD}' nicht gefunden. Schritt übersprungen."
    else
        if [ "$DRY_RUN" = true ]; then
            log_message "WARN" "[DRY RUN] Befehl würde im Verzeichnis '${PRODUCT_DEPOT_PATH}' ausgeführt: ${CREATE_DRIVER_LINKS_CMD}"
        else
            if [ "$COUNT_SUCCESS" -gt 0 ]; then
                log_message "INFO" "Wechsle temporär in das Arbeitsverzeichnis: ${PRODUCT_DEPOT_PATH}"
                pushd "${PRODUCT_DEPOT_PATH}" > /dev/null
                log_message "INFO" "Führe '${CREATE_DRIVER_LINKS_CMD}' aus..."
                python3 "$CREATE_DRIVER_LINKS_CMD" # Expliziter Aufruf mit python3 zur Sicherheit
                popd > /dev/null
                log_message "INFO" "Zurück zum ursprünglichen Arbeitsverzeichnis."

                log_message "OK" "Treiber-Pakete erstellt."
                log_message "WARN" "WICHTIG: Pakete jetzt mit folgendem Befehl installieren:"
                log_message "WARN" "sudo ${OPSI_PKG_MGR_CMD} -i /var/lib/opsi/workbench/${PRODUCT_ID}*.opsi"
            else
                log_message "WARN" "Keine Treiber kopiert, Paket-Erstellung übersprungen."
            fi
        fi
    fi
fi

print_header "Schritt 8/9: WinPE-Boot-Image aktualisieren"
if [ "$DRY_RUN" = true ]; then
    log_message "WARN" "[DRY RUN] Dieser Schritt wird im echten Lauf interaktiv sein."
else
    if [ "$COUNT_SUCCESS" -gt 0 ]; then
        read -p "Sollen die Treiber für '${PRODUCT_ID}' jetzt in das WinPE-Image integriert werden? (kann dauern) [j/N]: " -n 1 -r; echo
        if [[ $REPLY =~ ^[Jj]$ ]]; then
            log_message "INFO" "Führe '${OPSI_SETUP_CMD} --update-winpe' aus..."
            "$OPSI_SETUP_CMD" --update-winpe
            log_message "OK" "WinPE-Image wurde aktualisiert."
        else
            log_message "WARN" "Schritt übersprungen. Manuell ausführen mit: sudo ${OPSI_SETUP_CMD} --update-winpe"
        fi
    else
        log_message "WARN" "Keine Treiber kopiert, WinPE-Update nicht notwendig."
    fi
fi

# --- QUELLE LÖSCHEN (OPTIONAL) ---
print_header "Schritt 9/9: Bereinigung der Quell-Verzeichnisse"
if [ "$DRY_RUN" = true ]; then
    log_message "WARN" "[DRY RUN] Löschoperationen werden übersprungen."
elif [ "$DELETE_SOURCE_AFTER_COPY" = "on_success" ]; then
    log_message "INFO" "Modus 'on_success': Lösche alle erfolgreich kopierten Quellverzeichnisse..."
    if [ ${#SUCCESSFULLY_COPIED_SOURCES[@]} -gt 0 ]; then
        for dir_to_delete in "${SUCCESSFULLY_COPIED_SOURCES[@]}"; do
            if [ -d "$dir_to_delete" ]; then
                 log_message "INFO" "[LÖSCHE QUELLE] '${dir_to_delete}/'"
                 rm -rf "${dir_to_delete}"
            fi
        done
        log_message "OK" "${#SUCCESSFULLY_COPIED_SOURCES[@]} Quellverzeichnisse erfolgreich gelöscht."
    else
        log_message "INFO" "Keine Verzeichnisse zum Löschen vorhanden."
    fi
elif [ "$DELETE_SOURCE_AFTER_COPY" = "per_file" ]; then
    log_message "INFO" "Modus 'per_file': Löschen wurde bereits während des Kopiervorgangs durchgeführt."
else # never
    log_message "INFO" "Modus 'never': Keine Quellverzeichnisse werden gelöscht."
fi

# --- ABSCHLIESSENDE ZUSAMMENFASSUNG ---
print_header "Zusammenfassung des Durchlaufs für '${PRODUCT_ID}'"
log_message "OK" "Skriptdurchlauf beendet."
log_message "INFO" "Insgesamt .inf-Dateien im Quellverzeichnis: ${TOTAL_INF_FILES}"
log_message "OK" "Anzahl der eindeutigen Treiber-Verzeichnisse, die kopiert wurden: ${COUNT_SUCCESS}"
log_message "INFO" "Alle Details finden Sie in der Log-Datei: ${LOG_FILE}"
if [[ $DRY_RUN = true ]]; then
    log_message "WARN" "Dies war ein Trockenlauf. Es wurden keine Änderungen am System vorgenommen."
fi
echo ""

if [[ $DEBUG = true ]]; then
    set +x
fi

exit 0

