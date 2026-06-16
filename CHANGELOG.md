# Changelog - Skripts-and-tools

All notable changes to this repository are documented in this file.

## [1.1.0] - 2026-06-16
### Added
- Standardized directory layout under scripts/ (AD, System, User, Network).
- Unified JSON configuration system (Import-Environment.ps1 and environment.json.example).
- Consolidated L-Kennung query scripts into Get-LKennungUser.ps1.
- Consolidated L-Kennung report export utilities into New-LKennungReport.ps1.
- Consolidated DNS server setup and repair tools into Setup-DnsServer.ps1.
- Migrated legacy scripts to rchive/.

### Changed
- Standardized script naming to standard PowerShell Verb-Noun PascalCase format.
- Moved win11-hardening package to general repository and created link junction.
- Reconstructed commit history for Copy-ADUser.ps1, Set-LKennungPassword.ps1, and Enhanced-AD-User-Management.ps1.

### Removed
- Obsolete root scripts and duplicate WinPE builders (migrated to Schul-OPSI).
