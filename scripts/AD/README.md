# Active Directory Administration Scripts

This directory contains scripts for managing Active Directory users, groups, and permissions.

## Scripts

- **Copy-ADUser.ps1**: Copies an Active Directory user account, including properties and group memberships.
- **Enhanced-AD-User-Management.ps1**: Comprehensive interactive tool for AD user provisioning, CSV imports, property synchronization, and L-Kennung exports.
- **Export-ADClientsToCSV.ps1**: Queries and exports AD computer objects for specific OUs to CSV.
- **Get-DeponieUserReasons.ps1**: Inspects disabled/deponie user accounts to determine reasons for inactivity.
- **Get-LKennungUser.ps1**: Consolidated read-only AD query tool for email addresses, password expiry, and FINDUS groups.
- **Manage-ADUsers.ps1**: Core utility to manage general AD user attributes and group memberships.
- **Manage-NLUsers.ps1**: Special manager for NL-prefix user accounts.
- **New-ADUserFromCSV.ps1**: Creates bulk AD users from a UTF-8 delimited CSV template.
- **New-LKennungReport.ps1**: Consolidates reporting tools to generate CSV, Excel (with group color coding), and HTML D3.js visualization.
- **Reset-ADUserPassword.ps1**: Utility script to reset user passwords and enable accounts.
- **Set-ADUserProperties.ps1**: Configures user account properties based on templates.
- **Set-LKennungLastLogon.ps1**: Updates last logon timestamps for tracking purposes.
- **Set-LKennungPassword.ps1**: Secure write utility for resetting L-Kennung passwords.
- **Sync-ADTree.ps1**: Triggers replication across all domain controllers in the forest.
