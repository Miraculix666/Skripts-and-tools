#!/bin/bash
################################################################################
# OPSI Repository Configuration Script
# Configures both UIB and official OPSI package sources
# Author: OPSI Administrator
# Date: $(date)
################################################################################

set -e

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# Configuration variables
REPO_DIR="/etc/opsi/package-updater.repos.d"
LOG_FILE="/var/log/opsi-repo-config.log"

# Function to print colored output
print_status() {
    echo -e "${BLUE}[INFO]${NC} $1"
}

print_success() {
    echo -e "${GREEN}[SUCCESS]${NC} $1"
}

print_error() {
    echo -e "${RED}[ERROR]${NC} $1" >&2
}

print_warning() {
    echo -e "${YELLOW}[WARNING]${NC} $1"
}

# Check if running as root
if [[ $EUID -ne 0 ]]; then
   print_error "This script must be run as root"
   exit 1
fi

# Create repository directory if it doesn't exist
if [ ! -d "$REPO_DIR" ]; then
    print_status "Creating repository configuration directory: $REPO_DIR"
    mkdir -p "$REPO_DIR"
fi

# Log function
log_action() {
    echo "[$(date +'%Y-%m-%d %H:%M:%S')] $1" >> "$LOG_FILE"
}

print_status "Starting OPSI Repository Configuration..."
log_action "Starting OPSI Repository Configuration"

################################################################################
# 1. Configure Official UIB OPSI Repository
################################################################################

print_status "Configuring Official UIB OPSI Repository..."

cat > "$REPO_DIR/uib_official.repo" << 'EOF'
[repository_uib_official]
; Official UIB OPSI package repository
; Source: https://www.uib.de/en/

active = true

; Main official UIB repository
baseURL = http://opsipackages.43.opsi.org/stable

; Directory structure for different OPSI versions and products
dirs = opsi4.3/products/localboot,opsi4.3/products/netboot,opsi4.3/linux,opsi4.3/windows

; Standard product installation behavior
autoInstall = true
autoUpdate = true
autoSetup = false
onlyDownload = false

; Error handling
ignoreErrors = true

; Timeout and proxy settings (uncomment if needed)
; timeout = 300
; proxy = http://proxy.example.com:8080

description = Official UIB OPSI Package Repository (Stable)
EOF

print_success "Created UIB official repository configuration"
log_action "UIB official repository configured: $REPO_DIR/uib_official.repo"

################################################################################
# 2. Configure Official OPSI Packages Repository
################################################################################

print_status "Configuring Official OPSI Packages Repository..."

cat > "$REPO_DIR/opsi_packages_official.repo" << 'EOF'
[repository_opsi_packages]
; Official OPSI Packages Repository
; Direct access to opsipackages.43.opsi.org

active = true

baseURL = http://opsipackages.43.opsi.org/stable

; Complete directory structure
dirs = opsi4.3/products/localboot,opsi4.3/products/netboot

; Installation behavior
autoInstall = true
autoUpdate = true
autoSetup = false
onlyDownload = false

; Error handling
ignoreErrors = true

description = Official OPSI Packages Repository

EOF

print_success "Created OPSI packages official repository configuration"
log_action "OPSI packages official repository configured: $REPO_DIR/opsi_packages_official.repo"

################################################################################
# 3. Configure OPSI4Institutes (o4i) Repository (Community)
################################################################################

print_status "Configuring OPSI4Institutes Community Repository..."

cat > "$REPO_DIR/o4i_public.repo" << 'EOF'
[repository_o4i_public]
; OPSI4Institutes Public Repository
; Community-driven open source OPSI packages
; Website: https://o4i.org/

active = true

; Multiple mirror options available
baseURL = https://repo.o4i.org/public

; Public branch packages
dirs = /

; Installation behavior
autoInstall = true
autoUpdate = true
autoSetup = false
onlyDownload = false

; Error handling
ignoreErrors = true

description = OPSI4Institutes (o4i) Public Repository

EOF

print_success "Created OPSI4Institutes public repository configuration"
log_action "OPSI4Institutes repository configured: $REPO_DIR/o4i_public.repo"

################################################################################
# 4. Configure KIT/SCC Repository (Optional - German University)
################################################################################

print_status "Configuring KIT/SCC Repository (optional)..."

cat > "$REPO_DIR/kit_scc_repository.repo" << 'EOF'
[repository_kit_scc]
; KIT Steinbuch Centre for Computing OPSI Repository
; University repository for German academic institutions
; Reference: https://www.scc.kit.edu/en/services/10786.php

active = false
; Set to true if you want to use this repository

baseURL = http://opsi.scc.kit.edu/repository

dirs = /

autoInstall = false
autoUpdate = true
autoSetup = false
onlyDownload = false

ignoreErrors = true

description = KIT/SCC OPSI Repository (German Universities)

EOF

print_success "Created KIT/SCC repository configuration (disabled by default)"
log_action "KIT/SCC repository configured: $REPO_DIR/kit_scc_repository.repo"

################################################################################
# 5. Display All Repository Files
################################################################################

print_status "Displaying repository configurations..."
echo ""
echo -e "${BLUE}Repository Configuration Files:${NC}"
ls -la "$REPO_DIR"/*.repo

################################################################################
# 6. Validate OPSI Configuration
################################################################################

print_status "Validating OPSI configuration..."

# Check if opsi-package-updater is available
if command -v opsi-package-updater &> /dev/null; then
    print_status "opsi-package-updater found"

    # List active repositories
    echo ""
    echo -e "${BLUE}Active Repositories:${NC}"
    opsi-package-updater list --active-repos || print_warning "Could not list active repositories"

else
    print_warning "opsi-package-updater not found in PATH"
fi

################################################################################
# 7. Summary and Next Steps
################################################################################

print_success "OPSI Repository Configuration Completed!"
echo ""
echo -e "${BLUE}Summary:${NC}"
echo "=================================="
echo "✓ UIB Official Repository configured"
echo "✓ OPSI Packages Official Repository configured"
echo "✓ OPSI4Institutes (o4i) Public Repository configured"
echo "✓ KIT/SCC Repository configured (disabled)"
echo ""
echo -e "${BLUE}Next Steps:${NC}"
echo "=================================="
echo "1. Review repository configurations:"
echo "   cat $REPO_DIR/*.repo"
echo ""
echo "2. Update OPSI package index:"
echo "   opsi-package-updater list --products"
echo ""
echo "3. Download and install packages:"
echo "   opsi-package-updater install"
echo ""
echo "4. Set up automatic updates (cron job):"
echo "   # Add to crontab: 0 5 * * * /usr/bin/opsi-package-updater update"
echo ""
echo "5. Enable/disable repositories as needed:"
echo "   # Edit files in: $REPO_DIR"
echo "   # Change 'active = true/false'"
echo ""
echo -e "${BLUE}Log file:${NC} $LOG_FILE"
log_action "OPSI Repository Configuration Completed Successfully"
