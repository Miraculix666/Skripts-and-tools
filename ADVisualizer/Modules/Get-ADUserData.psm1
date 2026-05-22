# Module for retrieving and processing AD user data
function Get-ADUserDataEnhanced {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$SearchPattern,

        [Parameter()]
        [string]$SearchBase = "DC=domain,DC=local",

        [Parameter()]
        [string[]]$ExcludeOUs = @("Benutzer", "Disabled Users")
    )

    try {
        Write-Verbose "Retrieving AD users matching pattern '$SearchPattern' in $SearchBase..."

        $getADUserParams = @{
            Filter     = "SamAccountName -like '$SearchPattern'"
            Properties = "SamAccountName", "Name", "MemberOf", "DistinguishedName", "Comment", "Description"
        }
        if ($SearchBase) { $getADUserParams.SearchBase = $SearchBase }

        $users = Get-ADUser @getADUserParams

        if (-not $users) {
            Write-Warning "No users found matching pattern '$SearchPattern'"
            return $null
        }

        $userData = @()
        $matrixTable = @()
        $groupColors = @{}
        $colorIndex = 20  # Consistent with the monolithic script
        $allGroups = @{}
        $allOUs = @{}

        foreach ($user in $users) {
            Write-Verbose "Processing user: $($user.SamAccountName)"

            # Extract OU hierarchy
            $ouPath = ($user.DistinguishedName -split ',') | Where-Object { $_ -like 'OU=*' } | ForEach-Object { ($_ -split '=')[1] }
            $ouString = $ouPath -join '/'
            if (-not $ouPath) { $ouPath = @("No OU"); $ouString = "No OU" }

            $numericPrefix = if ($ouPath[0] -match '^\d{2,3}') { $Matches[0] } else { "999" }

            # Skip excluded OUs
            if ($ExcludeOUs -contains $ouPath[0]) {
                Write-Verbose "Skipping user in excluded OU: $($ouPath[0])"
                continue
            }

            # Track unique OUs
            if (-not $allOUs.ContainsKey($ouString)) {
                $allOUs[$ouString] = @{
                    Path          = $ouPath
                    Level         = $ouPath.Count
                    NumericPrefix = $numericPrefix
                }
            }

            # Get and sort groups
            $groups = $user.MemberOf | ForEach-Object {
                try {
                    (Get-ADGroup $_).Name
                }
                catch {
                    Write-Warning "Could not resolve group for user $($user.SamAccountName): $_"
                    return "Unknown Group"
                }
            } | Sort-Object

            # Create matrix format entry
            $matrixEntry = [PSCustomObject]@{
                NumericPrefix  = $numericPrefix
                OUPath         = $ouString
                Benutzer       = $user.Name
                SamAccountName = $user.SamAccountName
                Description    = $user.Description
            }

            # Add group columns to matrix entry
            foreach ($group in $groups) {
                if (-not $allGroups.ContainsKey($group)) {
                    $allGroups[$group] = $colorIndex
                    $groupColors[$group] = $colorIndex
                    $colorIndex++
                    if ($colorIndex -gt 56) { $colorIndex = 20 }
                }
                Add-Member -InputObject $matrixEntry -NotePropertyName $group -NotePropertyValue "X"

                $userData += [PSCustomObject]@{
                    SortPrefix     = $numericPrefix
                    OU             = $ouString
                    Benutzer       = $user.Name
                    SamAccountName = $user.SamAccountName
                    Gruppe         = $group
                    Kommentar      = $user.Comment
                    Description    = $user.Description
                    ColorIndex     = $allGroups[$group]
                }
            }
            $matrixTable += $matrixEntry
        }

        return @{
            DetailedData = $userData
            MatrixData   = $matrixTable
            Groups       = $allGroups
            OUs          = $allOUs
        }
    }
    catch {
        throw "Error retrieving AD user data: $_"
    }
}

Export-ModuleMember -Function Get-ADUserDataEnhanced
