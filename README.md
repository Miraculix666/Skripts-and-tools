# General Scripts and Tools Repository

A centralized collection of systems administration, Active Directory, network, and remote desktop utility scripts.

## Directory Structure

- **scripts/**: Main scripts grouped by domain
  - **AD/**: Active Directory user, properties, and L-Kennung management
  - **System/**: Software inventory, DNS management, and local system automation
  - **User/**: RDP managers and session launchers
  - **Network/**: Wake-on-LAN and networking tools
- **win11-hardening/**: OPSI package for Windows 11 client security configuration
- **archive/**: Archive for legacy or abandoned utility scripts

## Setup & Environment Configuration

General scripts utilize a central JSON-based configuration loader to manage environmental variables (e.g. AD domains, SMTP servers, default output directories) and credentials securely.

1. Copy environment.json.example to environment.json in the root of the repository.
2. Edit environment.json to match your local network and system configuration.
3. Call Import-Environment.ps1 in your scripts to load the configuration.

*Note: environment.json is ignored by Git to prevent leaking credentials.*
