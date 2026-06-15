<#
.SYNOPSIS
    Queries Active Directory for L-Kennung users, resolving Email, Password Expiration, Last Logon, and FINDUS group membership in a single pass.

.DESCRIPTION
    This script consolidates and replaces L-Kennung-Email.ps1, L-Kennung-PW-expiring-Date.ps1, and L-Kennung-FINDUS.ps1.
    It queries specified OUs (default: OU 81 and OU 82) for active L-Kennung users (L110* / L114*), collects their metadata,
    and exports detailed reports.

.PARAMETER OUs
    An array of Distinguished Names (DNs) of the OUs to search. If not specified, defaults to OUs 81 and 82.

.PARAMETER Filter
    The search filter to find matching users. Defaults to active users starting with L110 or L114.

.PARAMETER ExpiryThresholdWeeks
    The number of weeks of inactivity before a user is marked as expired. Defaults to 39 weeks.

.PARAMETER OutputDir
    The directory where reports will be saved. Defaults to 'C:\Daten'.

.NOTES
    Version: 1.0
    Autor: Systems Administration
    Compatibility: PowerShell 5.1+
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string[]]$OUs,

    [Parameter(Mandatory = $false)]
    [string]$Filter = "Enabled -eq $true -and (Name -like 'L110*' -or Name -like 'L114*')",

    [Parameter(Mandatory = $false)]
    [int]$ExpiryThresholdWeeks = 39,

    [Parameter(Mandatory = $false)]
    [string]$OutputDir = "C:\Daten"
)

# Load environment configuration if available in script directory or parent directories
$envLoader = Join-Path $PSScriptRoot "..\..\Import-Environment.ps1"
if (Test-Path $envLoader) {
    . $envLoader
    Write-Verbose "Loaded environment variables from loader."
}

# Apply default OUs if not specified
if (-not $OUs) {
    if ($global:Env -and $global:Env.OUs) {
        $OUs = $global:Env.OUs
    } else {
        $OUs = @(
            "OU=81,OU=Polizei-NRW-PB-PE-2012,DC=polizei,DC=nrw,DC=de",
            "OU=82,OU=Polizei-NRW-PB-PE-2012,DC=polizei,DC=nrw,DC=de"
        )
    }
}

# Apply default OutputDir from environment if set
if ($global:Env -and $global:Env.DefaultOutputDir) {
    $OutputDir = $global:Env.DefaultOutputDir
}

# Ensure output directory exists
if (-not (Test-Path $OutputDir -PathType Container)) {
    try {
        New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
    } catch {
        Write-Error "Failed to create output directory '$OutputDir': $_"
        return
    }
}

# Define file paths
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$mainReportCsv = Join-Path $OutputDir "L-Kennung_User_Query_Report_$timestamp.csv"
$allSamFile = Join-Path $OutputDir "All_SAMAccountNames.txt"
$expiredSamFile = Join-Path $OutputDir "Expired_Users_SAM.txt"
$expiredNoEmailFile = Join-Path $OutputDir "Expired_Users_NoEmail.txt"

# Clear or initialize output files
$null = New-Item -Path $allSamFile -ItemType File -Force
$null = New-Item -Path $expiredSamFile -ItemType File -Force
$null = New-Item -Path $expiredNoEmailFile -ItemType File -Force

# Calculate threshold date for inactivity
$thresholdDate = (Get-Date).AddDays(-$ExpiryThresholdWeeks * 7)

# Retrieve users across all target OUs
$users = [System.Collections.Generic.List[PSObject]]::new()
foreach ($ouPath in $OUs) {
    Write-Verbose "Searching OU: $ouPath"
    try {
        $foundUsers = Get-ADUser -Filter $Filter -SearchBase $ouPath -Properties LastLogonDate, EmailAddress, msDS-UserPasswordExpiryTimeComputed, MemberOf
        if ($foundUsers) {
            $users.AddRange($foundUsers)
        }
    } catch {
        Write-Warning "Could not query OU '$ouPath': $_"
    }
}

if ($users.Count -eq 0) {
    Write-Host "No users matching '$Filter' found in target OUs." -ForegroundColor Yellow
    return
}

Write-Host "Found $($users.Count) users. Processing details..." -ForegroundColor Cyan

$processedUsers = foreach ($user in $users) {
    # Determine password expiration date
    $expiryTime = $user."msDS-UserPasswordExpiryTimeComputed"
    $expiryDateStr = "N/A"
    $isPasswordExpired = $false
    if ($expiryTime -ne $null -and $expiryTime -gt 0 -and $expiryTime -ne 9223372036854775807) {
        try {
            $expiryDate = [datetime]::FromFileTime($expiryTime)
            $expiryDateStr = $expiryDate.ToString("yyyy-MM-dd HH:mm:ss")
            if ($expiryDate -lt (Get-Date)) {
                $isPasswordExpired = $true
            }
        } catch {
            $expiryDateStr = "Error parsing expiry"
        }
    } elseif ($expiryTime -eq 9223372036854775807) {
        $expiryDateStr = "Never Expired"
    }

    # Extract last logon date and format
    $lastLogon = $user.LastLogonDate
    $lastLogonStr = "N/A"
    $isInactive = $false
    if ($lastLogon -ne $null) {
        $lastLogonStr = $lastLogon.ToString("yyyy-MM-dd HH:mm:ss")
        if ($lastLogon -lt $thresholdDate) {
            $isInactive = $true
            $lastLogonStr = "xxxxx $lastLogonStr"
        }
    } else {
        $isInactive = $true
        $lastLogonStr = "xxxxx N/A"
    }

    # Parse MemberOf to find FINDUS groups
    $findusGroups = @()
    if ($user.MemberOf) {
        foreach ($groupDn in $user.MemberOf) {
            # Extract CN from Group DN: CN=Group_FINDUS_X,OU=...,DC=...
            if ($groupDn -match '^CN=([^,]+)') {
                $groupName = $Matches[1]
                if ($groupName -like "*FINDUS*") {
                    $findusGroups += $groupName
                }
            }
        }
    }
    $findusGroupsStr = $findusGroups -join ", "

    # Write SAMAccountName to total list
    $user.SamAccountName | Add-Content -Path $allSamFile

    # Check expiration/inactivity criteria
    if ($isInactive -or $isPasswordExpired) {
        $user.SamAccountName | Add-Content -Path $expiredSamFile
        if (-not $user.EmailAddress) {
            $user.SamAccountName | Add-Content -Path $expiredNoEmailFile
        }
    }

    [PSCustomObject]@{
        Name                 = $user.Name
        SamAccountName       = $user.SamAccountName
        EmailAddress         = $user.EmailAddress
        LastLogonDate        = $lastLogonStr
        PasswordExpiryDate   = $expiryDateStr
        FindusGroups         = $findusGroupsStr
        DistinguishedName    = $user.DistinguishedName
    }
}

# Sort by name
$sortedUsers = $processedUsers | Sort-Object Name

# Save main CSV report
$sortedUsers | Export-Csv -Path $mainReportCsv -NoTypeInformation -Delimiter ";" -Encoding UTF8

Write-Host "Query completed successfully." -ForegroundColor Green
Write-Host "Reports Saved to: $OutputDir"
Write-Host "  - Comprehensive CSV: $(Split-Path $mainReportCsv -Leaf)"
Write-Host "  - All SAM List:      $(Split-Path $allSamFile -Leaf)"
Write-Host "  - Expired SAM List:  $(Split-Path $expiredSamFile -Leaf)"
Write-Host "  - Expired No-Email:  $(Split-Path $expiredNoEmailFile -Leaf)"

# Show summary table to screen
$sortedUsers | Select-Object Name, SamAccountName, EmailAddress, LastLogonDate, PasswordExpiryDate, FindusGroups | Format-Table -AutoSize
