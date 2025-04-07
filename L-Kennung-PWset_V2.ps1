# Requires Active Directory module and RSAT tools
Import-Module ActiveDirectory

# Set debug parameters
$DebugPreference = 'Continue'
$ErrorActionPreference = 'Stop'

# Define search parameters
$searchPatterns = @('L110*', 'L114*')
$ouNames = @('81', '82')

# Function to find OUs recursively
function Find-TargetOUs {
    param([string[]]$Names)
    
    $domain = Get-ADDomain
    Write-Debug "Searching in domain: $($domain.DNSRoot)"
    
    $foundOUs = foreach ($name in $Names) {
        Get-ADOrganizationalUnit -Filter "Name -eq '$name'" `
            -SearchBase $domain.DistinguishedName `
            -SearchScope Subtree `
            -Properties DistinguishedName
    }
    
    if (-not $foundOUs) {
        throw "Target OUs ($($Names -join ', ')) not found in domain structure!"
    }
    
    return $foundOUs
}

try {
    # Find target OUs
    $targetOUs = Find-TargetOUs -Names $ouNames
    Write-Debug "Found OUs:`n$($targetOUs | Format-Table Name, DistinguishedName -AutoSize | Out-String)"
    
    # Search for users in each OU
    $users = foreach ($ou in $targetOUs) {
        foreach ($pattern in $searchPatterns) {
            Write-Debug "Searching in OU $($ou.DistinguishedName) for $pattern"
            
            Get-ADUser -LDAPFilter "(sAMAccountName=$pattern)" `
                -SearchBase $ou.DistinguishedName `
                -Properties * `
                -SearchScope Subtree
        }
    }
    
    if (-not $users) {
        Write-Host "No users found matching the criteria!" -ForegroundColor Yellow
        exit
    }
    
    # Display results
    Write-Host "Found $($users.Count) matching users:"
    $users | Format-Table Name, SamAccountName, DistinguishedName -AutoSize
    
    # Export results
    $csvPath = "C:\Daten\User_Report_$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
    $users | Export-Csv -Path $csvPath -Delimiter ';' -Encoding UTF8 -NoTypeInformation
    Write-Host "Report exported to: $csvPath" -ForegroundColor Green

}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Error details: $($_.ScriptStackTrace)" -ForegroundColor DarkGray
    exit 1
}
