#!/bin/bash

# ==============================================================================
# opsi-driver-sorter.sh V31.0 - Robuste Analyse ohne DriverVer-Zwang
# Erstellt für: OPSI-Server (Debian-basiert)
# Autor: PS-Coding
# Erstelldatum: 23.09.2025
# ==============================================================================
#
# Zweck: Analysiert, versioniert und sortiert Windows-Treiber für automatische
#        Treiberintegration in Windows-Netboot-Produkte (z.B. win10-x64).
#
# Funktionen:
# - Modus-Auswahl: Alle Treiber oder nur audit-basiert benötigte
# - Hardware-Audit von OPSI-Clients mit korrekten API-Aufrufen
# - Robuste INF-Analyse: Verarbeitet auch Treiber ohne explizite 'DriverVer'-Angabe
# - Versionserkennung und Datenbankmanagement
# - Strukturiertes Kopieren nach Gerätemanager-Kategorien
# - OPSI-Paket-Erstellung und WinPE-Integration
# - Universelle Pfadstruktur (produktübergreifend)
# - 8.3-konforme Pfade und Dateinamen (max. 240 Zeichen)
#
# Verwendung:
# ./opsi-driver-sorter.sh [-p NETBOOT_PRODUKT] [-h] [-v] [-q] [-d]
# 
# Parameter:
# -p NETBOOT_PRODUKT: OPSI-Netboot-Produkt-ID (Standard: win10-x64)
# -h: Zeigt diese Hilfe an
# -v: Verbose-Modus (Standard)
# -q: Quiet-Modus
# -d: Dry-Run-Modus (keine Änderungen)
#
# Quellen:
# - AI-entdeckt: OPSI-Dokumentation (https://docs.opsi.org/opsi-docs-de/4.3/)
# - Benutzer-bereitgestellt: Analyse fehlgeschlagener Versionen
# - Benutzer-bereitgestellt: Universelle Pfadstruktur-Anforderungen
# - Benutzer-bereitgestellt: Gerätemanager-basierte Struktur
#
# ==============================================================================

# ==============================================================================
# KONFIGURATION UND GLOBALE VARIABLEN
# ==============================================================================

# Standard-Konfiguration
readonly DEFAULT_NETBOOT_PRODUCT="win11-x64"
readonly SCRIPT_VERSION="31.0"
readonly SCRIPT_NAME="$(basename "$0")"

# Universelle Pfad-Konfiguration (produktübergreifend gemäß Vorgaben)
readonly BASE_DEPOT_PATH="/var/lib/opsi/depot"
readonly LOG_BASE_PATH="/var/log/opsi"
readonly WORKBENCH_PATH="/var/lib/opsi/workbench"

# Universelle Datenverzeichnisse (nicht produktspezifisch)
readonly UNIVERSAL_SOURCE_PATH="/var/lib/opsi/depot/drivers_source"
readonly UNIVERSAL_CACHE_FILE="/var/lib/opsi/depot/audit_hw_ids.cache"
readonly UNIVERSAL_DB_FILE="/var/lib/opsi/depot/driver.db"

# OPSI-Befehle (korrekte Pfade)
readonly OPSI_PYTHON_PATH="/usr/bin/opsi-python"
readonly OPSI_ADMIN_PATH="/usr/bin/opsi-admin"
readonly OPSI_SET_RIGHTS_PATH="/usr/bin/opsi-set-rights"
readonly OPSI_SETUP_PATH="/usr/bin/opsi-setup"
readonly OPSI_PACKAGE_MANAGER_PATH="/usr/bin/opsi-package-manager"

# Variablen für Laufzeit-Konfiguration
NETBOOT_PRODUCT=""
TARGET_PATH=""
LOG_FILE=""
VERBOSE_MODE=true
DRY_RUN_MODE=false
INTEGRATION_MODE=""

# Status-Tracking-Arrays
declare -A LATEST_DRIVERS
declare -A PROCESSED_DIRS
declare -A NEEDED_HW_IDS
ERROR_COUNT=0
WARNING_COUNT=0
COPY_COUNT=0

# Deutsche Lokalisierung und Kodierung sicherstellen
export LANG=de_DE.UTF-8
export LC_ALL=de_DE.UTF-8

# Gerätemanager-basierte Kategorien (gekürzt für 8.3-Kompatibilität)
declare -A DEVICE_CATEGORIES=(
    ["System"]="system"
    ["Audio"]="audio" 
    ["Bluetooth"]="bt"
    ["Chipset"]="chipset"
    ["Display"]="display"
    ["Graphics"]="graphic"
    ["Network"]="network"
    ["LAN"]="lan"
    ["WLAN"]="wlan"
    ["USB"]="usb"
    ["Storage"]="storage"
    ["RAID"]="raid"
    ["Security"]="security"
    ["Sensor"]="sensor"
    ["Monitor"]="monitor"
    ["Printer"]="printer"
    ["Scanner"]="scanner"
    ["Camera"]="camera"
    ["Modem"]="modem"
    ["Input"]="input"
    ["HID"]="hid"
    ["Firmware"]="firmware"
    ["Software"]="software"
)

# ==============================================================================
# HILFS- UND LOGGING-FUNKTIONEN
# ==============================================================================

# Hilfsfunktion anzeigen (vollständig deutsch)
show_help() {
    cat << EOF
$SCRIPT_NAME V$SCRIPT_VERSION - OPSI-Treiber-Sortierungsskript

VERWENDUNG:
    $SCRIPT_NAME [OPTIONEN]

OPTIONEN:
    -p NETBOOT_PRODUKT  OPSI-Netboot-Produkt-ID verarbeiten (Standard: $DEFAULT_NETBOOT_PRODUCT)
    -h                  Diese Hilfe anzeigen
    -v                  Verbose-Modus (Standard: aktiviert)
    -q                  Quiet-Modus (deaktiviert Verbose-Ausgabe)
    -d                  Dry-Run-Modus (keine Änderungen durchführen)

BEISPIELE:
    sudo $SCRIPT_NAME                       # Standard win10-x64 verarbeiten
    sudo $SCRIPT_NAME -p win11-x64          # Windows 11 x64 Netboot-Produkt
    sudo $SCRIPT_NAME -p server2022-x64      # Windows Server 2022 x64 Netboot-Produkt
    sudo $SCRIPT_NAME -d                      # Testlauf ohne Änderungen

BESCHREIBUNG:
    Dieses Skript analysiert Windows-Treiber aus dem universellen Quellverzeichnis
    ($UNIVERSAL_SOURCE_PATH), identifiziert die neuesten Versionen 
    basierend auf Hardware-Audits oder verarbeitet alle verfügbaren Treiber. 
    Auch Treiber ohne explizite 'DriverVer'-Angabe werden verarbeitet, solange
    gültige Hardware-Informationen oder System-Klassen vorhanden sind. 
    Die Treiber werden in eine strukturierte Zielhierarchie kopiert und 
    OPSI-Pakete für das angegebene Netboot-Produkt erstellt.

UNIVERSELLE PFADSTRUKTUR:
    Quelle:         $UNIVERSAL_SOURCE_PATH
    Cache:          $UNIVERSAL_CACHE_FILE
    Datenbank:      $UNIVERSAL_DB_FILE
    Ziel:           /var/lib/opsi/depot/[NETBOOT_PRODUKT]/drivers/drivers/

GERÄTEMANAGER-STRUKTUR:
    Treiber werden nach Windows Gerätemanager-Kategorien sortiert:
    - system, audio, bt, chipset, display, graphic
    - network, lan, wlan, usb, storage, raid
    - security, sensor, monitor, etc.

PFAD-OPTIMIERUNGEN:
    - 8.3-konforme Datei- und Verzeichnisnamen
    - Maximale Pfadlänge von 240 Zeichen wird angestrebt
    - Keine Leerzeichen oder Sonderzeichen in Pfaden
    - Automatische Sanitization aller Pfadkomponenten

VORAUSSETZUNGEN:
    - OPSI-Server mit funktionierender Installation
    - Zugriff auf $UNIVERSAL_SOURCE_PATH
    - Ausreichende Berechtigungen für OPSI-Befehle
    - Gültiges Netboot-Produkt in $BASE_DEPOT_PATH

WEITERE INFORMATIONEN:
    Siehe OPSI-Dokumentation: https://docs.opsi.org/opsi-docs-de/4.3/

EOF
}

# Erweiterte Logging-Funktion mit Zeitstempel und Farbcodierung (vollständig deutsch)
log_message() {
    local level="$1"
    local message="$2"
    local timestamp
    local color_code=""
    local reset_code="\033[0m"
    local plain_message
    
    timestamp=$(date '+%d.%m.%Y %H:%M:%S')
    plain_message="[$level] $message"
    
    # Farbcodes für verschiedene Log-Level
    case "$level" in
        "INFO")  color_code="\033[1;34m" ;;  # Blau
        "OK")    color_code="\033[1;32m" ;;  # Grün
        "WARN")  color_code="\033[1;33m" ;;  # Gelb
        "ERROR") color_code="\033[1;31m" ;;  # Rot
        "DEBUG") color_code="\033[1;35m" ;;  # Magenta
        *)       color_code="" ;;
    esac
    
    # Ausgabe auf Konsole mit Farbe (nur wenn Verbose-Modus aktiviert oder kritische Meldung)
    if [[ "$VERBOSE_MODE" == true ]] || [[ "$level" == "ERROR" ]] || [[ "$level" == "WARN" ]]; then
        echo -e "${color_code}${plain_message}${reset_code}"
    fi
    
    # Ausgabe in Log-Datei ohne Farbcodes
    if [[ -n "$LOG_FILE" ]]; then
        echo "[$timestamp] $plain_message" >> "$LOG_FILE"
    fi
    
    # Fehler- und Warnungszähler
    case "$level" in
        "ERROR") ((ERROR_COUNT++)) ;;
        "WARN")  ((WARNING_COUNT++)) ;;
    esac
}

# Header für Phasen ausgeben (vollständig deutsch)
print_header() {
    local message="$1"
    log_message "INFO" "=================================================================="
    log_message "OK" "$message"
    log_message "INFO" "=================================================================="
}

# Validierung der Systemvoraussetzungen (vollständig deutsch)
validate_prerequisites() {
    log_message "INFO" "Überprüfe Systemvoraussetzungen..."
    
    # OPSI-Befehle verfügbar?
    local required_commands=("$OPSI_ADMIN_PATH" "$OPSI_SET_RIGHTS_PATH" "$OPSI_PYTHON_PATH")
    for cmd in "${required_commands[@]}"; do
        if [[ ! -x "$cmd" ]]; then
            log_message "ERROR" "Erforderlicher Befehl nicht gefunden oder nicht ausführbar: $cmd"
            return 1
        fi
    done
    
    # Basis-Verzeichnisse verfügbar?
    local required_dirs=("$BASE_DEPOT_PATH" "$LOG_BASE_PATH")
    for dir in "${required_dirs[@]}"; do
        if [[ ! -d "$dir" ]]; then
            log_message "ERROR" "Erforderliches Verzeichnis nicht gefunden: $dir"
            return 1
        fi
        if [[ ! -w "$dir" ]]; then
            log_message "ERROR" "Keine Schreibberechtigung für: $dir"
            return 1
        fi
    done
    
    # Universelles Quellverzeichnis erstellen falls nicht vorhanden
    if [[ ! -d "$UNIVERSAL_SOURCE_PATH" ]]; then
        log_message "WARN" "Universelles Quellverzeichnis nicht gefunden: $UNIVERSAL_SOURCE_PATH"
        if [[ "$DRY_RUN_MODE" == false ]]; then
            mkdir -p "$UNIVERSAL_SOURCE_PATH" || {
                log_message "ERROR" "Konnte universelles Quellverzeichnis nicht erstellen: $UNIVERSAL_SOURCE_PATH"
                return 1
            }
            log_message "OK" "Universelles Quellverzeichnis erstellt: $UNIVERSAL_SOURCE_PATH"
        fi
    fi
    
    log_message "OK" "Systemvoraussetzungen erfüllt."
    return 0
}

# Netboot-Produkt validieren (vollständig deutsch)
validate_netboot_product() {
    local product_path="$BASE_DEPOT_PATH/$NETBOOT_PRODUCT"
    
    if [[ ! -d "$product_path" ]]; then
        log_message "ERROR" "Netboot-Produkt-Verzeichnis nicht gefunden: $product_path"
        log_message "ERROR" "Verfügbare Produkte:"
        ls -1 "$BASE_DEPOT_PATH" 2>/dev/null | grep -E "^(win|server)" | head -10 | while read -r product; do
            log_message "INFO" "  - $product"
        done
        return 1
    fi
    
    log_message "OK" "Netboot-Produkt validiert: $NETBOOT_PRODUCT"
    return 0
}

# ==============================================================================
# GERÄTEMANAGER-STRUKTUR UND KATEGORISIERUNG
# ==============================================================================

# Gerätekategorie aus INF-Datei bestimmen (nach Windows Gerätemanager-Struktur)
detect_device_category() {
    local inf_content="$1"
    local class_line device_class
    
    # Class= Zeile extrahieren
    class_line=$(echo "$inf_content" | grep -i "^Class[[:space:]]*=" | head -1)
    device_class=$(echo "$class_line" | cut -d'=' -f2 | tr -d ' "\r\n' | tr '[:upper:]' '[:lower:]')
    
    # Mapping auf Gerätemanager-Kategorien
    case "$device_class" in
        "system"|"computer"|"processor"|"systemdevices")
            echo "system"
            ;;
        "sound"|"media"|"audioendpoint")
            echo "audio"
            ;;
        "bluetooth"|"bluetoothradios")
            echo "bt"
            ;;
        "chipset"|"smbus"|"dmacontroller")
            echo "chipset"
            ;;
        "display"|"displayadapters")
            echo "display"
            ;;
        "net"|"network"|"networkadapters")
            echo "network"
            ;;
        "usb"|"universalserialbus")
            echo "usb"
            ;;
        "hdc"|"diskdrive"|"scsiadapter"|"storagecontrollers")
            echo "storage"
            ;;
        "security"|"tpm"|"securitydevices")
            echo "security"
            ;;
        "monitor"|"monitors")
            echo "monitor"
            ;;
        "printer"|"printqueue")
            echo "printer"
            ;;
        "camera"|"imagingdevices")
            echo "camera"
            ;;
        "modem"|"modems")
            echo "modem"
            ;;
        "hid"|"humaninterfacedevices"|"keyboard"|"mouse")
            echo "input"
            ;;
        "firmware"|"systemfirmware")
            echo "firmware"
            ;;
        "softwarecomponent"|"softwaredevices")
            echo "software"
            ;;
        *)
            # Fallback: Versuche aus Description oder anderen Hinweisen zu erraten
            if echo "$inf_content" | grep -qi -E "(graphic|video|vga|display)"; then
                echo "graphic"
            elif echo "$inf_content" | grep -qi -E "(network|ethernet|lan|nic)"; then
                echo "lan"
            elif echo "$inf_content" | grep -qi -E "(wireless|wifi|wlan|802\.11)"; then
                echo "wlan"
            elif echo "$inf_content" | grep -qi -E "(raid|scsi|sata|nvme)"; then
                echo "raid"
            elif echo "$inf_content" | grep -qi -E "(sensor|thermal|temperature)"; then
                echo "sensor"
            else
                echo "system"  # Default fallback
            fi
            ;;
    esac
}

# 8.3-konforme Verzeichnisnamen generieren
sanitize_directory_name() {
    local name="$1"
    local max_length="${2:-8}"
    
    # Kleinbuchstaben, nur alphanumerisch
    name=$(echo "$name" | tr '[:upper:]' '[:lower:]' | sed 's/[^a-z0-9]//g')
    
    # Länge begrenzen
    if [[ ${#name} -gt $max_length ]]; then
        name="${name:0:$max_length}"
    fi
    
    echo "$name"
}

# ==============================================================================
# PARAMETER-VERARBEITUNG UND INITIALISIERUNG
# ==============================================================================

# Kommandozeilenparameter verarbeiten (vollständig deutsch)
process_arguments() {
    while getopts "p:hvqd" opt; do
        case $opt in
            p)
                NETBOOT_PRODUCT="$OPTARG"
                ;;
            h)
                show_help
                exit 0
                ;;
            v)
                VERBOSE_MODE=true
                ;;
            q)
                VERBOSE_MODE=false
                ;;
            d)
                DRY_RUN_MODE=true
                ;;
            \?)
                log_message "ERROR" "Ungültige Option: -$OPTARG"
                show_help
                exit 1
                ;;
        esac
    done
    
    # Standardwerte setzen
    if [[ -z "$NETBOOT_PRODUCT" ]]; then
        NETBOOT_PRODUCT="$DEFAULT_NETBOOT_PRODUCT"
    fi
}

# Pfade und Log-Datei initialisieren (vollständig deutsch)
initialize_paths() {
    # Produktspezifisches Zielverzeichnis gemäß OPSI-Doku
    TARGET_PATH="$BASE_DEPOT_PATH/$NETBOOT_PRODUCT/drivers/drivers"
    
    # Log-Datei erstellen mit sicherem Namen
    local timestamp
    timestamp=$(date '+%Y-%m-%d_%H-%M-%S')
    local log_filename
    log_filename="driver_sort_${NETBOOT_PRODUCT}_${timestamp}.log"
    LOG_FILE="$LOG_BASE_PATH/$log_filename"
    
    # Netboot-Produkt validieren
    validate_netboot_product || return 1
    
    # Universelles Quellverzeichnis validieren
    if [[ ! -d "$UNIVERSAL_SOURCE_PATH" ]]; then
        log_message "ERROR" "Universelles Quellverzeichnis nicht gefunden: $UNIVERSAL_SOURCE_PATH"
        return 1
    fi
    
    # Zielverzeichnis erstellen falls nötig
    if [[ ! -d "$TARGET_PATH" ]] && [[ "$DRY_RUN_MODE" == false ]]; then
        mkdir -p "$TARGET_PATH" || {
            log_message "ERROR" "Konnte Zielverzeichnis nicht erstellen: $TARGET_PATH"
            return 1
        }
    fi
    
    log_message "OK" "Pfade initialisiert (universelle Struktur):"
    log_message "INFO" "  Universelle Quelle: $UNIVERSAL_SOURCE_PATH"
    log_message "INFO" "  Produktspezifisches Ziel: $TARGET_PATH"
    log_message "INFO" "  Universelle Datenbank: $UNIVERSAL_DB_FILE"
    log_message "INFO" "  Universeller Cache: $UNIVERSAL_CACHE_FILE"
    log_message "INFO" "  Log-Datei: $LOG_FILE"
    
    return 0
}

# ==============================================================================
# MODUS-AUSWAHL UND AUDIT-FUNKTIONEN
# ==============================================================================

# Benutzer nach Integrationsmodus fragen (vollständig deutsch)
select_integration_mode() {
    print_header "Phase 1/5: Auswahl der Integrationsmethode"
    
    echo "Wählen Sie den Integrationsmodus für '$NETBOOT_PRODUCT':"
    echo "1) Alle gefundenen Treiber verarbeiten (Standard-Integration)"
    echo "2) Nur von OPSI-Hardware-Audit benötigte Treiber (Audit-basierte Integration)"
    echo ""
    
    local choice
    read -p "Ihre Wahl [1]: " choice
    choice=${choice:-1}
    
    case $choice in
        1)
            INTEGRATION_MODE="all"
            log_message "OK" "Standard-Integration: Alle gefundenen Treiber werden verarbeitet."
            ;;
        2)
            INTEGRATION_MODE="audit"
            log_message "OK" "Audit-basierte Integration: Nur benötigte Treiber werden verarbeitet."
            ;;
        *)
            log_message "WARN" "Ungültige Auswahl. Verwende Standard-Integration."
            INTEGRATION_MODE="all"
            ;;
    esac
}

# Hardware-IDs von OPSI-Clients sammeln (universeller Cache) (vollständig deutsch)
collect_hardware_audit() {
    print_header "Phase 2/5: Hardware-Audit wird durchgeführt"
    
    local force_audit=false
    
    # Prüfen ob universeller Cache existiert
    if [[ -f "$UNIVERSAL_CACHE_FILE" ]]; then
        local cache_age
        cache_age=$(find "$UNIVERSAL_CACHE_FILE" -mtime +7 2>/dev/null | wc -l)
        if [[ $cache_age -gt 0 ]]; then
            echo "Universeller Audit-Cache ist älter als 7 Tage."
        else
            echo "Aktueller universeller Audit-Cache gefunden (weniger als 7 Tage alt)."
        fi
        
        local use_cache
        read -p "Vorhandenen universellen Cache verwenden? [j/N]: " -n 1 -r use_cache
        echo
        if [[ ! $use_cache =~ ^[Jj]$ ]]; then
            force_audit=true
        fi
    else
        force_audit=true
    fi
    
    if [[ "$force_audit" == true ]]; then
        log_message "INFO" "Führe neues Hardware-Audit durch (universeller Cache)..."
        perform_hardware_audit "$UNIVERSAL_CACHE_FILE"
    else
        log_message "INFO" "Verwende vorhandenen universellen Audit-Cache: $UNIVERSAL_CACHE_FILE"
    fi
    
    # Cache-Datei laden
    if [[ ! -s "$UNIVERSAL_CACHE_FILE" ]]; then # -s prüft, ob die Datei größer als 0 ist
        log_message "ERROR" "Die Audit-Cache-Datei ist leer. Es wurden keine Hardware-IDs gefunden. Bitte führen Sie das Skript erneut aus und erzwingen Sie ein neues Audit."
        return 1
    fi
    
    log_message "INFO" "Lese benötigte Hardware-IDs aus dem Cache..."
    while read -r line; do
        NEEDED_HW_IDS["${line^^}"]=1
    done < "$UNIVERSAL_CACHE_FILE"
    
    local hw_count=${#NEEDED_HW_IDS[@]}
    log_message "OK" "Hardware-Audit abgeschlossen. $hw_count eindeutige Hardware-IDs aus dem Cache geladen."
    return 0
}

# Eigentliches Hardware-Audit durchführen (vollständig deutsch, korrekte OPSI-API-Aufrufe)
perform_hardware_audit() {
    local cache_file="$1"
    local temp_file
    temp_file=$(mktemp)
    local client_count=0
    local processed_count=0
    
    log_message "INFO" "Sammle Client-Liste vom OPSI-Server..."
    
    # Client-IDs abrufen (bewährte Methode)
    local clients
    if ! clients=$($OPSI_ADMIN_PATH method getClientIds_list 2>/dev/null); then
        log_message "ERROR" "Konnte Client-Liste nicht abrufen. OPSI-Admin-Zugriff prüfen."
        rm -f "$temp_file"
        return 1
    fi
    
    # Clients verarbeiten
    local client_ids
    mapfile -t client_ids < <(echo "$clients" | grep -o '"[^"]*"' | tr -d '"')
    client_count=${#client_ids[@]}
    
    log_message "INFO" "$client_count Clients gefunden. Starte Abfrage der Hardware-Daten..."
    
    local current_client_num=0
    for client_id in "${client_ids[@]}"; do
        ((current_client_num++))
        log_message "INFO" "Verarbeite Client $current_client_num/$client_count: $client_id"
        
        # Hardware-Daten für diesen Client abrufen (bewährte Methode)
        local hw_data
        hw_data=$($OPSI_ADMIN_PATH method hardware_getHashes "" "hardwareClass='pci' or hardwareClass='usb'" "$client_id" 2>/dev/null || true)
        
        if [[ -n "$hw_data" ]]; then
            ((processed_count++))
            # PCI/USB-IDs extrahieren und normalisieren (korrigiertes Regex)
            echo "$hw_data" | grep -o -E '(pci|usb)-[0-9a-f_]+' | while read -r hw_id; do
                # Format: pci-8086_1234 -> PCI\VEN_8086&DEV_1234
                if [[ "$hw_id" =~ ^pci-([0-9a-f]{4})_([0-9a-f]{4}) ]]; then
                    echo "PCI\\VEN_${BASH_REMATCH[1]^^}&DEV_${BASH_REMATCH[2]^^}"
                elif [[ "$hw_id" =~ ^usb-([0-9a-f]{4})_([0-9a-f]{4}) ]]; then
                    echo "USB\\VID_${BASH_REMATCH[1]^^}&PID_${BASH_REMATCH[2]^^}"
                fi
            done >> "$temp_file"
        else
             log_message "WARN" "Keine Hardware-Daten für Client ${client_id} erhalten. Wird übersprungen."
        fi
    done
    
    # Eindeutige IDs sortieren und universellen Cache erstellen
    if [[ -s "$temp_file" ]]; then
        sort -u "$temp_file" > "$cache_file"
        log_message "OK" "Hardware-Audit abgeschlossen: $processed_count von $client_count Clients hatten verwertbare Daten."
    else
        log_message "WARN" "Keine Hardware-Daten gesammelt. Erstelle leeren universellen Cache."
        touch "$cache_file"
    fi
    rm -f "$temp_file"
}


# ==============================================================================
# TREIBER-ANALYSE-FUNKTIONEN (mit universeller DB und korrekter UTF-16LE Behandlung)
# ==============================================================================

# INF-Datei-Inhalt lesen und konvertieren (UTF-16LE zu UTF-8)
read_inf_content() {
    local inf_file="$1"
    iconv -f UTF-16LE -t UTF-8 "$inf_file" 2>/dev/null || cat "$inf_file" 2>/dev/null
}


# Universelle Treiber-Datenbank mit Dateisystem abgleichen (vollständig deutsch)
sync_driver_database() {
    print_header "Phase 3/5: Gleiche universelle Treiber-DB mit dem Dateisystem ab..."
    
    local temp_db
    temp_db=$(mktemp)
    local removed_count=0
    local total_count=0
    
    if [[ -f "$UNIVERSAL_DB_FILE" ]]; then
        while IFS='|' read -r inf_path checksum version category hw_ids; do
            ((total_count++))
            if [[ -f "$inf_path" ]]; then
                # Datei existiert noch, behalten
                echo "$inf_path|$checksum|$version|$category|$hw_ids" >> "$temp_db"
            else
                # Datei existiert nicht mehr, aus DB entfernen
                log_message "DEBUG" "Entferne verwaisten DB-Eintrag: $inf_path"
                ((removed_count++))
            fi
        done < "$UNIVERSAL_DB_FILE"
        
        if [[ -f "$temp_db" ]]; then
            mv "$temp_db" "$UNIVERSAL_DB_FILE"
        else
            # Wenn temp_db leer ist, wurde die Original-DB auch geleert
            >"$UNIVERSAL_DB_FILE"
        fi
    else
        touch "$UNIVERSAL_DB_FILE"
    fi
    
    log_message "OK" "Universelle Treiber-DB-Abgleich abgeschlossen. $removed_count von $total_count Einträgen entfernt."
}

# Neue oder geänderte Treiber analysieren (aus universeller Quelle) (vollständig deutsch)
analyze_drivers() {
    print_header "Phase 4/5: Analysiere neue/geänderte Treiber..."
    
    local file_count=0
    local processed_count=0
    local current_file=0
    
    # Alle INF-Dateien im universellen Quellverzeichnis finden
    log_message "INFO" "Suche nach .inf-Dateien in: $UNIVERSAL_SOURCE_PATH"
    
    # Stabile Schleife ohne Subshell
    local inf_files_list
    inf_files_list=$(mktemp)
    find "$UNIVERSAL_SOURCE_PATH" -iname "*.inf" -type f -print0 > "$inf_files_list"
    file_count=$(grep -c -z "" "$inf_files_list")

    if [[ $file_count -eq 0 ]]; then
        log_message "WARN" "Keine .inf-Dateien gefunden in: $UNIVERSAL_SOURCE_PATH"
        rm -f "$inf_files_list"
        return 1
    fi
    
    log_message "OK" "$file_count .inf-Dateien zur Analyse gefunden."
    
    # Jede INF-Datei verarbeiten
    while IFS= read -r -d $'\0' inf_file; do
        ((current_file++))
        
        log_message "INFO" "------------------------------------------------------------------"
        log_message "INFO" "Prüfe Datei ($current_file/$file_count): $inf_file"
        
        if analyze_single_inf_file "$inf_file" "$UNIVERSAL_DB_FILE"; then
            ((processed_count++))
        fi
        
        # Fortschrittsanzeige alle 100 Dateien
        if ((current_file % 100 == 0)); then
            log_message "INFO" "Fortschritt: $current_file von $file_count Dateien verarbeitet"
        fi
    done < "$inf_files_list"
    rm -f "$inf_files_list"
    
    log_message "OK" "Treiber-Analyse abgeschlossen: $processed_count von $file_count Dateien neu analysiert oder aktualisiert."
}

# Einzelne INF-Datei analysieren (korrekte UTF-16LE zu UTF-8 Konvertierung)
analyze_single_inf_file() {
    local inf_file="$1"
    local db_file="$2"
    local current_checksum
    local cached_checksum=""
    local driver_version=""
    local hw_ids=""
    local driver_date=""
    local device_category=""
    
    # Checksum berechnen
    current_checksum=$(md5sum "$inf_file" 2>/dev/null | cut -d' ' -f1)
    if [[ -z "$current_checksum" ]]; then
        log_message "WARN" "Konnte Checksum nicht berechnen für: $inf_file"
        return 1
    fi
    
    # Prüfen ob bereits in universeller DB und unverändert
    if [[ -f "$db_file" ]]; then
        cached_checksum=$(grep "^$inf_file|" "$db_file" 2>/dev/null | cut -d'|' -f2)
        if [[ "$cached_checksum" == "$current_checksum" ]]; then
            log_message "INFO" "Treiber ist unverändert (Cache-Treffer via Checksum). Überspringe Analyse."
            return 0
        fi
    fi
    
    log_message "INFO" "Neue/geänderte Datei. Führe Analyse durch..."
    
    # INF-Datei-Inhalt lesen und konvertieren (UTF-16LE zu UTF-8)
    local inf_content
    inf_content=$(read_inf_content "$inf_file")
    
    if [[ -z "$inf_content" ]]; then
        log_message "WARN" "Konnte INF-Datei nicht lesen: $inf_file"
        return 1
    fi
    
    # DriverVer-Zeile extrahieren
    local driver_ver_line
    driver_ver_line=$(echo "$inf_content" | grep -i "^DriverVer" | head -1)
    
    if [[ -z "$driver_ver_line" ]]; then
        log_message "WARN" "Keine 'DriverVer'-Zeile gefunden. Verwende Standard-Version 0.0.0.0 und fahre mit der Analyse fort."
        driver_date="01/01/1970"
        driver_version="0.0.0.0"
    else
        # Version und Datum parsen (korrekte Regex)
        if [[ "$driver_ver_line" =~ DriverVer[[:space:]]*=[[:space:]]*([^,]+),(.+) ]]; then
            driver_date="${BASH_REMATCH[1]}"
            driver_version="${BASH_REMATCH[2]}"
        else
            log_message "WARN" "Konnte 'DriverVer'-Zeile nicht korrekt parsen. Verwende Standard-Version."
            driver_version="0.0.0.0"
            driver_date="01/01/1970"
        fi
        log_message "INFO" "Version gefunden: $driver_date,$driver_version"
    fi
    
    # Hardware-IDs extrahieren
    hw_ids=$(extract_hardware_ids "$inf_content")
    
    # Gerätekategorie bestimmen (nach Gerätemanager-Struktur)
    device_category=$(detect_device_category "$inf_content")
    
    # Prüfen ob Treiber relevante Hardware-IDs hat oder System-Komponente ist
    if [[ -z "$hw_ids" ]]; then
        if [[ "$device_category" =~ ^(system|firmware|software)$ ]]; then
            hw_ids="SYSTEM_COMPONENT"
            log_message "INFO" "System-Komponente erkannt (Kategorie: $device_category)."
        else
            log_message "WARN" "Keine unterstützten HW-IDs und keine bekannte System-Kategorie gefunden. Datei wird ignoriert."
            return 1
        fi
    fi
    
    # Normalisierte Version für Vergleiche erstellen
    local normalized_version
    normalized_version=$(normalize_version "$driver_version" "$driver_date")
    
    # Universelle DB-Eintrag aktualisieren/hinzufügen
    if [[ "$DRY_RUN_MODE" == false ]]; then
        # Alten Eintrag entfernen falls vorhanden
        grep -v "^$inf_file|" "$db_file" > "$db_file.tmp" 2>/dev/null || touch "$db_file.tmp"
        
        # Neuen Eintrag hinzufügen: inf_pfad|checksum|normalized_version|category|hw_ids
        echo "$inf_file|$current_checksum|$normalized_version|$device_category|$hw_ids" >> "$db_file.tmp"
        mv "$db_file.tmp" "$db_file"
    fi
    
    log_message "OK" "Analyse erfolgreich. Universelle DB-Eintrag hinzugefügt/aktualisiert."
    return 0
}

# Hardware-IDs aus INF-Inhalt extrahieren (korrekte Regex)
extract_hardware_ids() {
    local inf_content="$1"
    local hw_ids=""
    
    # Eindeutige IDs sammeln
    local found_ids
    found_ids=$(echo "$inf_content" | grep -o -i -E 'PCI\\VEN_[0-9A-F]{4}&DEV_[0-9A-F]{4}|USB\\VID_[0-9A-F]{4}&PID_[0-9A-F]{4}' | sort -u)
    
    # Komma-separierte Liste erstellen
    hw_ids=$(echo "$found_ids" | tr '\n' ',' | sed 's/,$//')
    
    echo "$hw_ids"
}

# Treiber-Version normalisieren für Vergleiche (korrigierte Implementierung)
normalize_version() {
    local version="$1"
    local date="$2"
    local normalized=""
    
    # Datum zu numerischem Wert konvertieren (YYYYMMDD)
    local date_numeric="19700101"
    if [[ "$date" =~ ([0-9]{1,2})/([0-9]{1,2})/([0-9]{4}) ]]; then
        local month="${BASH_REMATCH[1]}"
        local day="${BASH_REMATCH[2]}"
        local year="${BASH_REMATCH[3]}"
        date_numeric=$(printf "%04d%02d%02d" "$year" "$month" "$day")
    fi
    
    # Version in numerisches Format bringen
    local version_parts
    IFS='.' read -ra version_parts <<< "$version"
    local version_numeric=""
    for part in "${version_parts[@]::4}"; do
        # Nur numerische Teile verwenden, auf max 5 Stellen begrenzt
        local numeric_part
        numeric_part=$(echo "$part" | grep -o '^[0-9]*' | head -c 5)
        version_numeric="${version_numeric}$(printf "%05d" "${numeric_part:-0}")"
    done
    
    # Kombinierter Wert: YYYYMMDD + Versionsnummer
    normalized="${date_numeric}${version_numeric}"
    echo "$normalized"
}

# ==============================================================================
# TREIBER-SELEKTION UND KOPIER-FUNKTIONEN (mit Gerätemanager-Struktur)
# ==============================================================================

# Neueste benötigte Treiber ermitteln (vollständig deutsch)
determine_latest_drivers() {
    print_header "Phase 5/5: Ermittle neueste benötigte Treiber..."
    
    local hw_id_filter=""
    
    if [[ ! -f "$UNIVERSAL_DB_FILE" ]]; then
        log_message "ERROR" "Universelle Treiber-Datenbank nicht gefunden: $UNIVERSAL_DB_FILE"
        return 1
    fi
    
    # Audit-Cache laden falls audit-basierte Integration
    if [[ "$INTEGRATION_MODE" == "audit" ]]; then
        if [[ ! -s "$UNIVERSAL_CACHE_FILE" ]]; then
            log_message "ERROR" "Universeller Audit-Cache ist leer. Audit-basierte Filterung nicht möglich."
            return 1
        fi
        
        # Hardware-IDs als Filter vorbereiten
        while read -r line; do
            NEEDED_HW_IDS["${line^^}"]=1
        done < "$UNIVERSAL_CACHE_FILE"
        log_message "INFO" "Audit-basierte Filterung aktiv. ${#NEEDED_HW_IDS[@]} Hardware-IDs geladen."
    else
        log_message "INFO" "Standard-Integration: Alle Treiber werden berücksichtigt."
    fi
    
    # Treiber nach Hardware-IDs gruppieren und neueste bestimmen
    declare -A hw_id_versions
    local line_count=0
    
    while IFS='|' read -r inf_path checksum version category hw_ids; do
        ((line_count++))
        [[ -z "$inf_path" ]] && continue
        
        # Hardware-IDs splitten
        IFS=',' read -ra hw_id_array <<< "$hw_ids"
        
        for hw_id in "${hw_id_array[@]}"; do
            [[ -z "$hw_id" ]] && continue
            local hw_id_upper=${hw_id^^}
            
            # Prüfen ob Hardware-ID benötigt wird (bei Audit-Modus)
            if [[ "$INTEGRATION_MODE" == "audit" ]] && [[ "$hw_id_upper" != "SYSTEM_COMPONENT" ]]; then
                if [[ -z "${NEEDED_HW_IDS[$hw_id_upper]}" ]]; then
                    continue  # Diese Hardware-ID wird nicht benötigt
                fi
            fi
            
            # Prüfen ob dies die neueste Version für diese Hardware-ID ist
            if [[ -z "${hw_id_versions[$hw_id_upper]}" ]] || [[ "$version" > "${hw_id_versions[$hw_id_upper]}" ]]; then
                hw_id_versions[$hw_id_upper]="$version"
                LATEST_DRIVERS["$hw_id_upper"]="$version;$inf_path;$category"
                log_message "DEBUG" "Neueste Version für $hw_id_upper: $version ($inf_path, Kategorie: $category)"
            fi
        done
    done < "$UNIVERSAL_DB_FILE"
    
    local selected_count=${#LATEST_DRIVERS[@]}
    log_message "OK" "Neueste Treiber ermittelt: $selected_count eindeutige Hardware-IDs/Komponenten."
    
    if [[ $selected_count -eq 0 ]]; then
        log_message "WARN" "Keine Treiber ausgewählt. Überprüfen Sie die Konfiguration."
        return 1
    fi
    
    return 0
}

# Ausgewählte Treiber kopieren (mit Gerätemanager-Struktur)
copy_selected_drivers() {
    print_header "Phase 6/9: Kopiere ausgewählte Treiber"
    
    local error_count=0
    local skipped_count=0
    
    for hw_id in "${!LATEST_DRIVERS[@]}"; do
        local driver_info="${LATEST_DRIVERS[$hw_id]}"
        local version inf_path category
        
        # Parse: version;inf_path;category
        IFS=';' read -r version inf_path category <<< "$driver_info"
        local source_dir
        source_dir=$(dirname "$inf_path")
        
        # Prüfen ob bereits verarbeitet (vermeidet Duplikate)
        if [[ -n "${PROCESSED_DIRS[$source_dir]}" ]]; then
            ((skipped_count++))
            continue
        fi
        
        # Sicheren Kategorie-Ordner-Namen generieren (8.3-konform)
        local safe_category_name
        safe_category_name=$(sanitize_directory_name "$category" 8)
        
        # Produktspezifisches Zielverzeichnis erstellen (z.B. .../man/pci/8086/1234)
        local vendor_id device_id
        if [[ "$hw_id" =~ VEN_([0-9A-F]{4})\&DEV_([0-9A-F]{4}) ]]; then
            vendor_id="${BASH_REMATCH[1]}"
            device_id="${BASH_REMATCH[2]}"
        elif [[ "$hw_id" =~ VID_([0-9A-F]{4})\&PID_([0-9A-F]{4}) ]]; then
            vendor_id="${BASH_REMATCH[1]}"
            device_id="${BASH_REMATCH[2]}"
        else # SYSTEM_COMPONENT
            vendor_id=$(sanitize_directory_name "$(basename "$source_dir")" 8)
            device_id="-"
        fi
        
        local target_dir="$TARGET_PATH/$safe_category_name/$vendor_id/$device_id"
        
        log_message "INFO" "Kopiere Treiber für $hw_id (Kategorie: $category):"
        log_message "INFO" "  Quelle: $source_dir"
        log_message "INFO" "  Sicheres Ziel: $target_dir"
        
        if [[ "$DRY_RUN_MODE" == false ]]; then
            # Zielverzeichnis erstellen
            if ! mkdir -p "$target_dir"; then
                log_message "ERROR" "Konnte Zielverzeichnis nicht erstellen: $target_dir"
                ((error_count++))
                continue
            fi
            
            # Treiber-Verzeichnis kopieren mit rsync
            if rsync -av --delete "$source_dir/" "$target_dir/" >> "$LOG_FILE" 2>&1; then
                log_message "OK" "Treiber erfolgreich kopiert."
                ((COPY_COUNT++))
                PROCESSED_DIRS["$source_dir"]="1"
            else
                log_message "ERROR" "Fehler beim Kopieren von: $source_dir"
                ((error_count++))
            fi
        else
            log_message "INFO" "[DRY-RUN] Würde kopieren: $source_dir -> $target_dir"
            ((COPY_COUNT++))
        fi
    done
    
    log_message "OK" "Kopierprozess abgeschlossen:"
    log_message "INFO" "  Erfolgreich kopiert: $COPY_COUNT"
    log_message "INFO" "  Übersprungen (Duplikate): $skipped_count"
    log_message "INFO" "  Fehler: $error_count"
    
    return $error_count
}

# ==============================================================================
# NACHBEARBEITUNGS-FUNKTIONEN
# ==============================================================================

# OPSI-Dateirechte setzen (vollständig deutsch)
fix_file_permissions() {
    print_header "Schritt 7/9: Dateirechte werden korrigiert"
    
    if [[ "$DRY_RUN_MODE" == true ]]; then
        log_message "WARN" "[DRY RUN] Rechte-Korrektur wird übersprungen."
        return 0
    fi
    
    if [[ $COPY_COUNT -gt 0 ]]; then
        log_message "INFO" "Führe '$OPSI_SET_RIGHTS_PATH' für '$NETBOOT_PRODUCT' aus..."
        
        if "$OPSI_SET_RIGHTS_PATH" "$BASE_DEPOT_PATH/$NETBOOT_PRODUCT"; then
            log_message "OK" "Dateirechte wurden erfolgreich gesetzt."
        else
            log_message "ERROR" "Fehler beim Setzen der Dateirechte."
            return 1
        fi
    else
        log_message "WARN" "Keine Treiber verarbeitet. Rechte-Korrektur übersprungen."
    fi
    
    return 0
}

# OPSI-Treiberpakete erstellen (vollständig deutsch, korrektes Arbeitsverzeichnis)
create_driver_packages() {
    print_header "Schritt 8/9: OPSI-Treiberpakete werden erstellt"
    
    local create_script="$BASE_DEPOT_PATH/$NETBOOT_PRODUCT/create_driver_links.py"
    
    if [[ ! -f "$create_script" ]]; then
        log_message "WARN" "create_driver_links.py nicht gefunden in: $create_script"
        log_message "WARN" "OPSI-Pakete können nicht automatisch erstellt werden."
        return 1
    fi
    
    if [[ "$DRY_RUN_MODE" == true ]]; then
        log_message "WARN" "[DRY RUN] Paket-Erstellung wird übersprungen."
        return 0
    fi
    
    if [[ $COPY_COUNT -gt 0 ]]; then
        log_message "INFO" "Wechsle temporär in das Arbeitsverzeichnis: $BASE_DEPOT_PATH/$NETBOOT_PRODUCT"
        
        if pushd "$BASE_DEPOT_PATH/$NETBOOT_PRODUCT" > /dev/null; then
            log_message "INFO" "Führe '$create_script' aus (mit opsi-python)..."
            
            # Verwende opsi-python statt normales python3 - KORREKTE METHODE
            if "$OPSI_PYTHON_PATH" "$create_script"; then
                log_message "OK" "Treiber-Pakete erfolgreich erstellt."
            else
                log_message "ERROR" "Fehler bei der Paket-Erstellung."
                popd > /dev/null
                return 1
            fi
            
            popd > /dev/null
            log_message "INFO" "Zurück zum ursprünglichen Arbeitsverzeichnis."
            
            log_message "WARN" "WICHTIG: Pakete jetzt mit folgendem Befehl installieren:"
            log_message "WARN" "sudo $OPSI_PACKAGE_MANAGER_PATH -i $WORKBENCH_PATH/$NETBOOT_PRODUCT*.opsi"
        else
            log_message "ERROR" "Konnte nicht in Arbeitsverzeichnis wechseln."
            return 1
        fi
    else
        log_message "WARN" "Keine Treiber verarbeitet. Paket-Erstellung übersprungen."
    fi
    
    return 0
}

# WinPE-Boot-Image aktualisieren (vollständig deutsch)
update_winpe_image() {
    print_header "Schritt 9/9: WinPE-Boot-Image aktualisieren"
    
    if [[ "$DRY_RUN_MODE" == true ]]; then
        log_message "WARN" "[DRY RUN] WinPE-Update wird übersprungen."
        return 0
    fi
    
    if [[ $COPY_COUNT -gt 0 ]]; then
        local update_winpe
        read -p "Sollen die Treiber für '$NETBOOT_PRODUCT' jetzt in das WinPE-Image integriert werden? (kann dauern) [j/N]: " -n 1 -r update_winpe
        echo
        
        if [[ $update_winpe =~ ^[Jj]$ ]]; then
            log_message "INFO" "Starte WinPE-Update..."
            
            if "$OPSI_SETUP_PATH" --update-winpe; then
                log_message "OK" "WinPE-Image erfolgreich aktualisiert."
            else
                log_message "ERROR" "Fehler beim WinPE-Update."
                return 1
            fi
        else
            log_message "WARN" "Schritt übersprungen. Manuell ausführen mit: sudo $OPSI_SETUP_PATH --update-winpe"
        fi
    else
        log_message "WARN" "Keine Treiber verarbeitet. WinPE-Update übersprungen."
    fi
    
    return 0
}


# ==============================================================================
# HAUPTPROGRAMM
# ==============================================================================

# Aufräum-Funktion für EXIT-Handler (vollständig deutsch)
cleanup() {
    local exit_code=$?
    
    if [[ $exit_code -eq 0 && $ERROR_COUNT -eq 0 ]]; then
        log_message "OK" "Skript erfolgreich abgeschlossen."
    else
        log_message "ERROR" "Skript mit Fehlern beendet (Exit-Code: $exit_code)."
    fi
    
    log_message "INFO" "Zusammenfassung für Netboot-Produkt '$NETBOOT_PRODUCT':"
    log_message "INFO" "  Verarbeitete Treiber: $COPY_COUNT"
    log_message "INFO" "  Warnungen: $WARNING_COUNT"
    log_message "INFO" "  Fehler: $ERROR_COUNT"
    log_message "INFO" "  Universelle Quelle: $UNIVERSAL_SOURCE_PATH"
    log_message "INFO" "  Universelle DB: $UNIVERSAL_DB_FILE"
    log_message "INFO" "  Universeller Cache: $UNIVERSAL_CACHE_FILE"
    log_message "INFO" "  Log-Datei: $LOG_FILE"
    
    # Temporäre Dateien sicher entfernen
    rm -f /tmp/hw_audit_$$ /tmp/driver_db_sync_$$
    
    exit $exit_code
}

# Main-Funktion (vollständig deutsch)
main() {
    # EXIT-Handler registrieren
    trap cleanup EXIT
    
    # Header ausgeben
    log_message "INFO" "Starte OPSI Treiber-Sortierungsskript V$SCRIPT_VERSION (Universelle Pfadstruktur)"
    log_message "INFO" "Verarbeite Netboot-Produkt: $NETBOOT_PRODUCT"
    
    if [[ "$DRY_RUN_MODE" == true ]]; then
        log_message "WARN" "DRY-RUN-MODUS: Keine Änderungen werden durchgeführt."
    fi
    
    # Schritt-für-Schritt-Verarbeitung
    validate_prerequisites || return 1
    
    # Modus-Auswahl
    select_integration_mode || return 1
    
    # Hardware-Audit (nur wenn audit-basiert)
    if [[ "$INTEGRATION_MODE" == "audit" ]]; then
        collect_hardware_audit || return 1
    fi
    
    # Treiber-Verarbeitung (universelle DB und Quelle)
    sync_driver_database || return 1
    analyze_drivers || return 1
    determine_latest_drivers || return 1
    copy_selected_drivers || return 1
    
    # Nachbearbeitung (produktspezifisch)
    fix_file_permissions || return 1
    create_driver_packages || return 1
    update_winpe_image || return 1
    
    return 0
}

# ==============================================================================
# SKRIPT-START
# ==============================================================================

# Prüfen ob als Root ausgeführt (vollständig deutsch)
if [[ $EUID -ne 0 ]]; then
    echo "FEHLER: Dieses Skript muss als Root ausgeführt werden."
    echo "Verwenden Sie: sudo $0 $*"
    exit 1
fi

# Parameter verarbeiten und Pfade initialisieren
process_arguments "$@"
initialize_paths || exit 1

# Hauptprogramm starten
main "$@"

