<#
.SYNOPSIS
    Queries the security event log of a domain controller for account lockout events (ID 4740).
.DESCRIPTION
    This script searches for Failure Audit security events on a domain controller, matching
    a specific username, and returns details such as time, domain controller, caller computer,
    and source network address.
.PARAMETER UserName
    The username to search for. Defaults to the current environment username.
.PARAMETER DomainController
    The Domain Controller (FQDN or NetBIOS name) to query. Defaults to the domain logon server.
.PARAMETER Days
    The number of days to search back. Defaults to 7 days.
.EXAMPLE
    .\Get-LogonErrorEvents.ps1 -UserName "john.doe" -DomainController "dc1.domain.local" -Days 3
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false, HelpMessage = "Username to search for lockout events.")]
    [string]$UserName = $env:USERNAME,

    [Parameter(Mandatory = $false, HelpMessage = "Domain Controller name or IP to query.")]
    [string]$DomainController = ($env:LOGONSERVER -replace '^\\+'),

    [Parameter(Mandatory = $false, HelpMessage = "Number of days to search back.")]
    [int]$Days = 7
)

$StartTime = (Get-Date).AddDays(-$Days)
$EndTime = Get-Date

if ([string]::IsNullOrEmpty($DomainController)) {
    Write-Error "No domain controller specified, and LOGONSERVER environment variable is not set."
    return
}

Write-Verbose "Querying Security log on $DomainController for user '$UserName' lockout events since $StartTime..."

try {
    Get-EventLog -ComputerName $DomainController -LogName Security -EntryType FailureAudit -After $StartTime -Before $EndTime |
    Where-Object {$_.EventID -eq 4740 -and $_.Message -like "*Account Name: $($UserName)*"} |
    Select-Object TimeGenerated, @{Name="DomainController";Expression={$_.MachineName}},
                  @{Name="LockedOutAccount";Expression={$_.Message -split "`n" | Where-Object {$_ -like "*Account Name*" } | ForEach-Object {$_.Split(":")[1].Trim()}}},
                  @{Name="CallerComputer";Expression={$_.Message -split "`n" | Where-Object {$_ -like "*Caller Computer Name*" } | ForEach-Object {$_.Split(":")[1].Trim()}}},
                  @{Name="SourceNetworkAddress";Expression={$_.Message -split "`n" | Where-Object {$_ -like "*Source Network Address*" } | ForEach-Object {$_.Split(":")[1].Trim()}}} |
    Sort-Object TimeGenerated -Descending
}
catch {
    Write-Error "Failed to query Security log on $DomainController. Details: $($_.Exception.Message)"
}
