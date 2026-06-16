# Windows 11 Hardening Package

This package manages Windows 11 local security hardening policies and configuration options.
Resides in the general scripts repository but is integrated transparently into OPSI deployment structures.

## Package Structure

- **CLIENT_DATA/**: Contains installation scripts and security policy registries.
  - Apply_Hardening.ps1: Script to apply CIS-benchmark security registry keys.
  - setup.opsi: OPSI installation control script.
- **OPSI/**: Contains OPSI package control files.
  - control: Package metadata and dependencies.
