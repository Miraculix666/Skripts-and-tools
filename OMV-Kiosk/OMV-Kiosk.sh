#!/bin/bash
# ==============================================================================
# Dateiname: OMV-Kiosk-Setup.sh
# Beschreibung: Installiert OpenMediaVault + lokalen XFCE-Kiosk +
#               Kopia Backup-Manager auf einem frischen Debian 12 System.
#
# Architektur:
#   ┌─────────────────────────────────────────────────────────────────────┐
#   │  Debian 12 (Bookworm) – Basis                                       │
#   │                                                                     │
#   │  ┌──────────────────────────────────────────────────────────────┐   │
#   │  │  OpenMediaVault 7                                            │   │
#   │  │  ├── WebUI   :80/:443  (Remote-Verwaltung, Storage-Config)  │   │
#   │  │  ├── ZFS Plugin        (Pools, Datasets, Snapshots)         │   │
#   │  │  └── SMB/NFS           (Netzwerkfreigaben)                  │   │
#   │  └──────────────────────────────────────────────────────────────┘   │
#   │                                                                     │
#   │  ┌──────────────────────────────────────────────────────────────┐   │
#   │  │  XFCE Desktop (Autologin → kioskuser)                       │   │
#   │  │  ├── Thunar            Dateimanager → /srv/dev-disk-by-*    │   │
#   │  │  ├── Kopia UI          Backup/Archiv Verwaltung             │   │
#   │  │  └── Chromium          OMV-WebUI lokal erreichbar           │   │
#   │  └──────────────────────────────────────────────────────────────┘   │
#   │                                                                     │
#   │  udisks2 → USB-Sticks erscheinen automatisch in Thunar             │
#   │  Kopia Dienst → läuft im Hintergrund, schreibt auf ZFS-Dataset     │
#   └─────────────────────────────────────────────────────────────────────┘
#
# Voraussetzungen:
#   - Frisches Debian 12 (Bookworm) – Minimalinstallation
#   - Root-Zugriff
#   - Internetverbindung
#
# Persistenz: Alle Installationen überleben OMV-Updates (Debian apt-Basis)
#
# Autor:     PS-Coding (AI-Assistent)
# Version:   1.0
# ==============================================================================

set -euo pipefail

# ==============================================================================
# Konfiguration – hier anpassen
# ==============================================================================

# --- Benutzer ---
KIOSK_USER="kioskuser"
KIOSK_PASS="kiosk123"           # Nur für LightDM-Autologin benötigt
ADMIN_USER="admin"              # OMV Web-Admin (wird separat gesetzt)

# --- Pfade ---
# OMV legt Datasets unter /srv/dev-disk-by-label-<NAME> ab
# Passe nach OMV-Einrichtung in der WebUI an
NAS_DATA_ROOT="/srv"            # OMV-Standard-Mountpunkt für Daten

# --- Kopia Backup ---
KOPIA_REPO_PATH="/srv/backup-repo"   # Wo Kopia sein Repository anlegt
KOPIA_PORT="51515"                   # Kopia Web-UI Port
# Weitere Kopia-Ziele (S3, SFTP) nach Installation in WebUI konfigurieren

# --- Filebrowser (optional, für einfachen Web-Zugriff zusätzlich zu OMV) ---
INSTALL_FILEBROWSER=1
FILEBROWSER_PORT="8080"
FILEBROWSER_DIR="/opt/filebrowser"

# --- Pakete ---
PKGS_DESKTOP="xfce4 thunar thunar-volman gvfs gvfs-backends udisks2 \
              dbus-x11 lightdm lightdm-gtk-greeter xfce4-terminal \
              tumbler ffmpegthumbnailer xfce4-notifyd"
PKGS_TOOLS="chromium curl wget git htop nmon smartmontools"

# --- Logging ---
LOG_FILE="/var/log/OMV-Kiosk-Setup_$(date +%Y%m%d_%H%M%S).log"
SILENT_MODE=0

# ==============================================================================
# Parameter
# ==============================================================================
while [[ "$#" -gt 0 ]]; do
    case $1 in
        --auto|-a)            SILENT_MODE=1 ;;
        --no-filebrowser)     INSTALL_FILEBROWSER=0 ;;
        --kiosk-user)         KIOSK_USER="$2"; shift ;;
        --data-root)          NAS_DATA_ROOT="$2"; shift ;;
        --kopia-repo)         KOPIA_REPO_PATH="$2"; shift ;;
        --help|-h)
            cat <<HELP
Verwendung: $0 [Optionen]

Optionen:
  --auto, -a              Unbeaufsichtigte Installation
  --no-filebrowser        Filebrowser nicht installieren
  --kiosk-user <USER>     Kiosk-Benutzername (Standard: kioskuser)
  --data-root <PFAD>      NAS-Datenpfad (Standard: /srv)
  --kopia-repo <PFAD>     Kopia-Repository-Pfad (Standard: /srv/backup-repo)
  --help, -h              Diese Hilfe

Ablauf:
  1. OpenMediaVault 7 installieren
  2. ZFS + omv-extras Plugin installieren
  3. XFCE Desktop + LightDM (Autologin)
  4. Thunar + udisks2 (USB-Automount)
  5. Kopia (Backup/Archiv-Manager)
  6. Filebrowser (optionaler Web-Zugriff)
  7. Chromium → OMV-WebUI

Nach Installation:
  OMV WebUI: http://<IP>      (admin / openmediavault)
  Kopia UI:  http://<IP>:${KOPIA_PORT}
  Kiosk:     XFCE öffnet automatisch mit Thunar + Kopia
HELP
            exit 0 ;;
        *) echo "Unbekannter Parameter: $1"; exit 1 ;;
    esac
    shift
done

# ==============================================================================
# Logging
# ==============================================================================
mkdir -p "$(dirname "$LOG_FILE")"
touch "$LOG_FILE"

if [[ $SILENT_MODE -eq 0 ]] && [[ -t 1 ]]; then
    R='\e[31m'; G='\e[32m'; Y='\e[33m'; B='\e[34m'; N='\e[0m'; BOLD='\e[1m'
else
    R=''; G=''; Y=''; B=''; N=''; BOLD=''
fi

log() {
    local LV="$1" MSG="$2"
    echo "[$(date +'%Y-%m-%d %H:%M:%S')] [${LV}] ${MSG}" >> "$LOG_FILE"
    case "$LV" in
        ERROR) echo -e "${R}[FEHLER]${N}   ${MSG}" ;;
        WARN)  echo -e "${Y}[WARNUNG]${N}  ${MSG}" ;;
        OK)    echo -e "${G}[OK]${N}       ${MSG}" ;;
        STEP)  echo -e "${B}${BOLD}[>>>>]${N}     ${MSG}" ;;
        *)     echo "           ${MSG}" ;;
    esac
}

die()  { log "ERROR" "$1"; log "ERROR" "Logfile: $LOG_FILE"; exit 1; }
ask()  {
    [[ $SILENT_MODE -eq 1 ]] && return 0
    read -rp "$1 [J/n]: " a; [[ "${a,,}" =~ ^(n|nein)$ ]] && return 1; return 0
}
run_q() {
    local D="$1"; shift
    if "$@" >> "$LOG_FILE" 2>&1; then log "OK" "$D"
    else die "$D fehlgeschlagen"; fi
}

# ==============================================================================
# Voraussetzungen
# ==============================================================================
check_prerequisites() {
    log "STEP" "Prüfe Voraussetzungen..."
    [[ "$EUID" -ne 0 ]] && die "Bitte als root ausführen"

    # Debian 12?
    if ! grep -q "bookworm\|VERSION_ID=\"12\"" /etc/os-release 2>/dev/null; then
        log "WARN" "Kein Debian 12 erkannt – OMV 7 benötigt Debian 12 (Bookworm)"
        ask "Trotzdem fortfahren?" || exit 0
    fi

    # OMV bereits installiert?
    if dpkg -l 2>/dev/null | grep -q "^ii  openmediavault "; then
        log "INFO" "OMV bereits installiert – überspringe OMV-Installation"
        OMV_INSTALLED=1
    else
        OMV_INSTALLED=0
    fi

    # Internet?
    ping -c 1 -W 5 packages.openmediavault.org &>/dev/null \
        || log "WARN" "OMV-Repository nicht erreichbar – Netzwerk prüfen"

    log "OK" "Voraussetzungen geprüft"
}

# ==============================================================================
# 1. OpenMediaVault installieren
# ==============================================================================
install_omv() {
    if [[ $OMV_INSTALLED -eq 1 ]]; then
        log "INFO" "OMV bereits installiert – überspringe"
        return 0
    fi

    log "STEP" "Installiere OpenMediaVault 7..."
    export DEBIAN_FRONTEND=noninteractive

    # Offizielle OMV-Installationsmethode für Debian
    # Quelle: https://docs.openmediavault.org/en/latest/installation/on_debian.html
    run_q "APT Update" apt-get update -qq
    run_q "Grundpakete" apt-get install -y -qq \
        gnupg curl apt-transport-https ca-certificates lsb-release

    # OMV GPG-Key + Repository
    curl -fsSL https://packages.openmediavault.org/public/archive.key \
        | gpg --dearmor -o /usr/share/keyrings/openmediavault.gpg \
        >> "$LOG_FILE" 2>&1 \
        || die "OMV GPG-Key konnte nicht hinzugefügt werden"

    cat > /etc/apt/sources.list.d/openmediavault.list <<EOF
deb [signed-by=/usr/share/keyrings/openmediavault.gpg] \
https://packages.openmediavault.org/public sandworm main
## Uncomment the following line to add software from the proposed repository.
# deb [signed-by=/usr/share/keyrings/openmediavault.gpg] \
# https://packages.openmediavault.org/public sandworm-proposed main
## Uncomment the following line to add software from the partner repository.
# deb [signed-by=/usr/share/keyrings/openmediavault.gpg] \
# https://packages.openmediavault.org/public sandworm partner
EOF

    run_q "APT Update (mit OMV-Repo)" apt-get update -qq
    run_q "OMV installieren" \
        apt-get install -y -qq \
        --no-install-recommends \
        openmediavault-plugin-developers \
        openmediavault

    # OMV initalisieren
    run_q "OMV Datenbank initialisieren" omv-confdbadm populate
    run_q "OMV Konfiguration anwenden"  omv-salt deploy run monit

    log "OK" "OpenMediaVault installiert"
    log "INFO" "WebUI: http://$(hostname -I | awk '{print $1}') | admin / openmediavault"
    log "WARN" "Admin-Passwort sofort nach Installation ändern!"
}

# ==============================================================================
# 2. omv-extras + ZFS Plugin
# ==============================================================================
install_zfs_plugin() {
    log "STEP" "Installiere omv-extras und ZFS-Plugin..."

    # omv-extras: Erweitert OMV um zusätzliche Plugins (inkl. ZFS)
    if ! dpkg -l 2>/dev/null | grep -q "^ii  omv-extras"; then
        wget -O /tmp/omv-extras.deb \
            https://github.com/OpenMediaVault-Plugin-Developers/packages/raw/master/openmediavault-omvextrasorg_latest_all.deb \
            >> "$LOG_FILE" 2>&1 \
            || die "omv-extras Download fehlgeschlagen"
        run_q "omv-extras installieren" dpkg -i /tmp/omv-extras.deb
        run_q "APT Update (omv-extras)" apt-get update -qq
        rm -f /tmp/omv-extras.deb
    else
        log "INFO" "omv-extras bereits installiert"
    fi

    # ZFS-Plugin
    if ! dpkg -l 2>/dev/null | grep -q "^ii  openmediavault-zfs"; then
        run_q "ZFS-Kernel-Module" apt-get install -y -qq zfsutils-linux
        run_q "OMV ZFS-Plugin"   apt-get install -y -qq openmediavault-zfs
    else
        log "INFO" "ZFS-Plugin bereits installiert"
    fi

    log "OK" "ZFS-Plugin installiert – Pools in OMV WebUI anlegen"
    log "INFO" "ZFS-Pools: Storage → ZFS → Pool hinzufügen"
}

# ==============================================================================
# 3. Desktop-Pakete
# ==============================================================================
install_desktop() {
    log "STEP" "Installiere Desktop-Umgebung..."
    export DEBIAN_FRONTEND=noninteractive

    run_q "APT Update"        apt-get update -qq
    run_q "Desktop-Pakete"    apt-get install -y -qq $PKGS_DESKTOP
    run_q "Tool-Pakete"       apt-get install -y -qq $PKGS_TOOLS

    log "OK" "Desktop installiert"
}

# ==============================================================================
# 4. Kiosk-Benutzer
# ==============================================================================
setup_kiosk_user() {
    log "STEP" "Richte Kiosk-Benutzer ein..."

    if ! id "$KIOSK_USER" &>/dev/null; then
        useradd -m -s /bin/bash "$KIOSK_USER" >> "$LOG_FILE" 2>&1
        echo "${KIOSK_USER}:${KIOSK_PASS}" | chpasswd >> "$LOG_FILE" 2>&1
        log "OK" "Benutzer $KIOSK_USER erstellt"
    else
        log "INFO" "Benutzer $KIOSK_USER existiert bereits"
    fi

    usermod -aG plugdev,disk,cdrom,audio,video,cdrom "$KIOSK_USER" \
        >> "$LOG_FILE" 2>&1 || true
}

# ==============================================================================
# 5. LightDM + XFCE konfigurieren
# ==============================================================================
configure_desktop() {
    log "STEP" "Konfiguriere Desktop (LightDM, XFCE, Thunar)..."

    # --- LightDM Autologin ---
    mkdir -p /etc/lightdm/lightdm.conf.d/
    cat > /etc/lightdm/lightdm.conf.d/50-kiosk.conf <<EOF
[Seat:*]
autologin-user=${KIOSK_USER}
autologin-user-timeout=0
user-session=xfce
greeter-session=lightdm-gtk-greeter
allow-user-switching=false
allow-guest=false
EOF

    # --- Kiosk-Verzeichnisse ---
    local CFG="/home/${KIOSK_USER}/.config"
    local AUTOSTART="${CFG}/autostart"
    local XCONF="${CFG}/xfce4/xfconf/xfce-perchannel-xml"
    local GTK3="${CFG}/gtk-3.0"

    mkdir -p "$AUTOSTART" "$XCONF" "$GTK3" \
             "${CFG}/Thunar"

    # --- Thunar: Lesezeichen (NAS-Root + USB) ---
    cat > "${GTK3}/bookmarks" <<EOF
file://${NAS_DATA_ROOT} NAS Daten
file:///media/${KIOSK_USER} USB-Laufwerke
EOF

    # --- Thunar: Einstellungen ---
    cat > "${CFG}/Thunar/thunarrc" <<EOF
[Configuration]
DefaultView=ThunarDetailsView
LastView=ThunarDetailsView
ShowHidden=FALSE
ShowSidePane=TRUE
SidePaneWidth=200
MiscShowDeleteAction=FALSE
MiscConfirmClose=FALSE
LastWindowMaximized=TRUE
LastWindowWidth=1280
LastWindowHeight=800
EOF

    # --- Autostart: Thunar-Daemon (USB-Erkennung) ---
    cat > "${AUTOSTART}/thunar-daemon.desktop" <<EOF
[Desktop Entry]
Type=Application
Name=Thunar Daemon
Exec=thunar --daemon
Terminal=false
Hidden=false
X-GNOME-Autostart-enabled=true
EOF

    # --- Autostart: Thunar öffnet NAS-Root ---
    cat > "${AUTOSTART}/thunar-nas.desktop" <<EOF
[Desktop Entry]
Type=Application
Name=NAS Dateimanager
Exec=bash -c 'sleep 2 && thunar ${NAS_DATA_ROOT}'
Terminal=false
Hidden=false
X-GNOME-Autostart-enabled=true
EOF

    # --- Autostart: Kopia UI ---
    cat > "${AUTOSTART}/kopia-ui.desktop" <<EOF
[Desktop Entry]
Type=Application
Name=Kopia Backup UI
Comment=Backup und Archiv Manager
Exec=bash -c 'sleep 4 && kopia-ui'
Terminal=false
Hidden=false
X-GNOME-Autostart-enabled=true
EOF

    # --- Autostart: Chromium → OMV WebUI ---
    cat > "${AUTOSTART}/omv-webui.desktop" <<EOF
[Desktop Entry]
Type=Application
Name=OMV Verwaltung
Comment=OpenMediaVault WebUI im Browser
Exec=bash -c 'sleep 5 && chromium --app=http://127.0.0.1 --new-window'
Terminal=false
Hidden=false
X-GNOME-Autostart-enabled=true
EOF

    # --- XFCE: Performance + Kiosk ---
    # Compositing aus (Legacy-HW)
    cat > "${XCONF}/xfwm4.xml" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<channel name="xfwm4" version="1.0">
  <property name="general" type="empty">
    <property name="use_compositing" type="bool" value="false"/>
  </property>
</channel>
EOF

    # Bildschirm niemals abschalten
    cat > "${XCONF}/xfce4-power-manager.xml" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<channel name="xfce4-power-manager" version="1.0">
  <property name="xfce4-power-manager" type="empty">
    <property name="dpms-enabled"       type="bool" value="false"/>
    <property name="blank-on-ac"        type="int"  value="0"/>
    <property name="presentation-mode"  type="bool" value="true"/>
    <property name="inactivity-on-ac"   type="uint" value="0"/>
  </property>
</channel>
EOF

    # Screensaver aus
    cat > "${XCONF}/xfce4-screensaver.xml" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<channel name="xfce4-screensaver" version="1.0">
  <property name="saver" type="empty">
    <property name="enabled"       type="bool" value="false"/>
    <property name="lock-enabled"  type="bool" value="false"/>
  </property>
</channel>
EOF

    # Tastenkombinationen
    cat > "${CFG}/xfce4/xfce4-keyboard-shortcuts.xml" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<channel name="xfce4-keyboard-shortcuts" version="1.0">
  <property name="commands" type="empty">
    <property name="custom" type="empty">
      <property name="&lt;Primary&gt;&lt;Alt&gt;t"
                type="string" value="xfce4-terminal"/>
      <property name="&lt;Primary&gt;&lt;Alt&gt;f"
                type="string" value="thunar ${NAS_DATA_ROOT}"/>
      <property name="&lt;Primary&gt;&lt;Alt&gt;k"
                type="string" value="kopia-ui"/>
      <property name="&lt;Primary&gt;&lt;Alt&gt;b"
                type="string" value="chromium --app=http://127.0.0.1"/>
    </property>
  </property>
</channel>
EOF

    # Berechtigungen
    chown -R "${KIOSK_USER}:${KIOSK_USER}" "/home/${KIOSK_USER}/.config"
    log "OK" "Desktop konfiguriert"
}

# ==============================================================================
# 6. udisks2 Policy (USB ohne Root-Passwort)
# ==============================================================================
configure_udisks() {
    log "STEP" "Konfiguriere udisks2 für USB-Automount..."

    mkdir -p /etc/polkit-1/rules.d/
    cat > /etc/polkit-1/rules.d/50-kiosk-udisks.rules <<EOF
polkit.addRule(function(action, subject) {
    var allow = [
        "org.freedesktop.udisks2.filesystem-mount",
        "org.freedesktop.udisks2.filesystem-unmount-others",
        "org.freedesktop.udisks2.eject-media",
        "org.freedesktop.udisks2.power-off-drive"
    ];
    if (allow.indexOf(action.id) >= 0 && subject.user === "${KIOSK_USER}") {
        return polkit.Result.YES;
    }
});
EOF

    log "OK" "udisks2: ${KIOSK_USER} darf USB-Laufwerke mounten"
    log "INFO" "USB erscheint unter: /media/${KIOSK_USER}/<Label>"
}

# ==============================================================================
# 7. Kopia installieren (Backup + Archiv)
# ==============================================================================
install_kopia() {
    log "STEP" "Installiere Kopia (Backup/Archiv-Manager)..."

    if command -v kopia &>/dev/null; then
        log "INFO" "Kopia bereits installiert: $(kopia --version 2>/dev/null || echo 'Version unbekannt')"
    else
        # Offizielles Kopia-Repository
        curl -fsSL https://kopia.io/signing-key \
            | gpg --dearmor -o /usr/share/keyrings/kopia-keyring.gpg \
            >> "$LOG_FILE" 2>&1 \
            || die "Kopia GPG-Key fehlgeschlagen"

        echo "deb [signed-by=/usr/share/keyrings/kopia-keyring.gpg] \
https://packages.kopia.io/apt/ stable main" \
            > /etc/apt/sources.list.d/kopia.list

        run_q "APT Update (Kopia)" apt-get update -qq
        # kopia     = CLI-Tool
        # kopia-ui  = Desktop-GUI (Electron)
        run_q "Kopia CLI + UI"     apt-get install -y -qq kopia kopia-ui
    fi

    # Kopia als Systemdienst (Web-UI auf Port KOPIA_PORT)
    # Kiosk-User als Repository-Owner
    cat > /etc/systemd/system/kopia-server.service <<EOF
[Unit]
Description=Kopia Repository Server (Web UI)
After=network.target

[Service]
Type=simple
User=${KIOSK_USER}
Group=${KIOSK_USER}
# Server startet die Web-UI + API
# Repository wird beim ersten Start über die UI angelegt
ExecStart=/usr/bin/kopia server start \
    --address=0.0.0.0:${KOPIA_PORT} \
    --server-username=admin \
    --server-password=kopia123 \
    --insecure \
    --without-password
Restart=on-failure
RestartSec=10s

[Install]
WantedBy=multi-user.target
EOF
    # Hinweis: --insecure nur für lokales Netz!
    # Für HTTPS: Zertifikat einrichten und --tls-cert-file / --tls-key-file setzen

    run_q "Systemd reload"         systemctl daemon-reload
    run_q "Kopia-Server aktivieren" systemctl enable kopia-server
    run_q "Kopia-Server starten"   systemctl start kopia-server

    mkdir -p "${KOPIA_REPO_PATH}"
    chown "${KIOSK_USER}:${KIOSK_USER}" "${KOPIA_REPO_PATH}"

    log "OK" "Kopia installiert"
    log "INFO" "Kopia WebUI: http://$(hostname -I | awk '{print $1}'):${KOPIA_PORT}"
    log "INFO" "Kopia Web-Login: admin / kopia123"
    log "INFO" "Repository-Pfad: ${KOPIA_REPO_PATH}"
    log "WARN" "Kopia-Passwörter nach Installation in der UI ändern!"
}

# ==============================================================================
# 8. Filebrowser (optional, zusätzlicher Web-Zugriff)
# ==============================================================================
install_filebrowser() {
    [[ $INSTALL_FILEBROWSER -eq 0 ]] && {
        log "INFO" "Filebrowser deaktiviert – übersprungen"
        return 0
    }

    log "STEP" "Installiere Filebrowser (zusätzlicher Web-Dateimanager)..."

    local ARCH; ARCH=$(uname -m)
    local FB_ARCH="amd64"
    [[ "$ARCH" == "aarch64" ]] && FB_ARCH="arm64"

    local FB_BIN="${FILEBROWSER_DIR}/filebrowser"
    mkdir -p "${FILEBROWSER_DIR}"

    if [[ ! -f "$FB_BIN" ]]; then
        # Neueste Version ermitteln
        local FB_VERSION
        FB_VERSION=$(curl -fsSL \
            https://api.github.com/repos/filebrowser/filebrowser/releases/latest \
            2>/dev/null | grep '"tag_name"' | grep -oP '\d+\.\d+\.\d+' | head -1 \
            || echo "2.27.0")

        local FB_URL="https://github.com/filebrowser/filebrowser/releases/download/v${FB_VERSION}/linux-${FB_ARCH}-filebrowser.tar.gz"
        log "INFO" "Lade Filebrowser v${FB_VERSION}..."

        curl -fsSL "$FB_URL" -o /tmp/fb.tar.gz >> "$LOG_FILE" 2>&1 \
            || { log "WARN" "Filebrowser-Download fehlgeschlagen – übersprungen"; return 0; }
        tar -xzf /tmp/fb.tar.gz -C "${FILEBROWSER_DIR}" filebrowser >> "$LOG_FILE" 2>&1
        chmod +x "$FB_BIN"
        rm -f /tmp/fb.tar.gz
    fi

    # Konfiguration: Zugriff aus LAN erlaubt
    cat > "${FILEBROWSER_DIR}/settings.json" <<EOF
{
  "port": ${FILEBROWSER_PORT},
  "address": "0.0.0.0",
  "database": "${FILEBROWSER_DIR}/filebrowser.db",
  "root": "${NAS_DATA_ROOT}",
  "log": "stdout",
  "auth": { "method": "noauth" },
  "branding": { "name": "NAS Dateizugriff", "color": "#1565C0" },
  "permissions": {
    "admin": false, "execute": false,
    "create": true,  "rename": true,
    "modify": true,  "delete": false,
    "share": false,  "download": true
  }
}
EOF

    "${FB_BIN}" config init \
        --config "${FILEBROWSER_DIR}/settings.json" >> "$LOG_FILE" 2>&1 || true

    cat > /etc/systemd/system/filebrowser.service <<EOF
[Unit]
Description=Filebrowser Web-Dateimanager
After=network.target

[Service]
Type=simple
User=root
ExecStart=${FB_BIN} \
    --config ${FILEBROWSER_DIR}/settings.json
Restart=on-failure
RestartSec=5s

[Install]
WantedBy=multi-user.target
EOF

    run_q "Systemd reload"          systemctl daemon-reload
    run_q "Filebrowser aktivieren"  systemctl enable filebrowser
    run_q "Filebrowser starten"     systemctl start filebrowser

    log "OK" "Filebrowser läuft auf Port ${FILEBROWSER_PORT}"
    log "WARN" "Ohne Passwort – nur im lokalen Netz nutzen!"
}

# ==============================================================================
# 9. Alle Dienste starten
# ==============================================================================
activate_services() {
    log "STEP" "Aktiviere Dienste..."
    run_q "Systemd reload"     systemctl daemon-reload
    run_q "LightDM aktivieren" systemctl enable lightdm
    run_q "LightDM starten"    systemctl restart lightdm
    log "OK" "Alle Dienste aktiv"
}

# ==============================================================================
# Zusammenfassung
# ==============================================================================
print_summary() {
    [[ $SILENT_MODE -eq 1 ]] && return 0
    local IP; IP=$(hostname -I 2>/dev/null | awk '{print $1}' || echo "?")

    echo ""
    echo -e "${BOLD}${G}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${N}"
    echo -e "${BOLD}${G}  INSTALLATION ABGESCHLOSSEN – OpenMediaVault Kiosk${N}"
    echo -e "${G}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${N}"
    echo ""
    echo -e "  ${BOLD}Remote-Zugriff:${N}"
    echo -e "  • OMV WebUI:     http://${IP}        (admin / openmediavault)"
    echo -e "  • Kopia WebUI:   http://${IP}:${KOPIA_PORT}  (admin / kopia123)"
    [[ $INSTALL_FILEBROWSER -eq 1 ]] && \
    echo -e "  • Filebrowser:   http://${IP}:${FILEBROWSER_PORT}  (kein Login)"
    echo ""
    echo -e "  ${BOLD}Lokal am Gerät (XFCE):${N}"
    echo -e "  • Thunar öffnet automatisch ${NAS_DATA_ROOT}"
    echo -e "  • Kopia UI startet automatisch"
    echo -e "  • USB erscheint in Thunar-Seitenleiste"
    echo ""
    echo -e "  ${BOLD}Tastenkürzel:${N}"
    echo -e "  • Ctrl+Alt+T  → Terminal"
    echo -e "  • Ctrl+Alt+F  → Thunar"
    echo -e "  • Ctrl+Alt+K  → Kopia UI"
    echo -e "  • Ctrl+Alt+B  → OMV WebUI in Chromium"
    echo ""
    echo -e "  ${BOLD}${Y}Nächste Schritte:${N}"
    echo -e "  1. OMV: admin-Passwort ändern"
    echo -e "  2. OMV: Storage → ZFS → Pool anlegen"
    echo -e "  3. Kopia: Repository auf ZFS-Dataset einrichten"
    echo -e "  4. Kopia: Backup-Jobs und Zeitplan konfigurieren"
    echo -e "  5. OMV: SMB/NFS-Freigaben für Netzwerkzugriff einrichten"
    echo ""
    echo -e "  ${BOLD}${R}Passwörter sofort ändern!${N}"
    echo -e "  OMV, Kopia und Filebrowser laufen initial ohne/mit"
    echo -e "  Standard-Passwort."
    echo ""
    echo -e "${G}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${N}"
}

# ==============================================================================
# Hauptprogramm
# ==============================================================================
main() {
    log "INFO" "════════════════════════════════════════"
    log "INFO" "OMV Kiosk Setup v1.0 | Silent: ${SILENT_MODE}"
    log "INFO" "════════════════════════════════════════"

    if [[ $SILENT_MODE -eq 0 ]]; then
        echo -e "${BOLD}"
        echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
        echo "  OpenMediaVault 7 + Kiosk-Desktop Setup"
        echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
        echo -e "${N}"
        echo "  Installiert auf Debian 12:"
        echo "  • OpenMediaVault 7  (NAS-Verwaltung)"
        echo "  • ZFS Plugin        (Pools, Snapshots)"
        echo "  • XFCE + LightDM    (Lokaler Desktop)"
        echo "  • Thunar + udisks2  (Dateimanager + USB)"
        echo "  • Kopia             (Backup/Archiv)"
        [[ $INSTALL_FILEBROWSER -eq 1 ]] && \
        echo "  • Filebrowser       (Web-Dateimanager)"
        echo ""
        ask "Installation starten?" || exit 0
    fi

    check_prerequisites
    install_omv
    install_zfs_plugin
    install_desktop
    setup_kiosk_user
    configure_desktop
    configure_udisks
    install_kopia
    install_filebrowser
    activate_services
    print_summary

    log "INFO" "Installation erfolgreich abgeschlossen."
}

main "$@"