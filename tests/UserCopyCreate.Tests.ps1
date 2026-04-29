<#
.SYNOPSIS
Pester tests for the Export-ADUsers function in UserCopyCreate.ps1.
#>

$ScriptPath = Join-Path -Path $PSScriptRoot -ChildPath "..\UserCopyCreate.ps1"

# Dot-source the script to load the functions.
# However, to avoid side-effects from the script body running, we might need to mock or isolate it.
# Since it's a procedural script, we will mock the necessary functions and set the script variables.

Describe "Export-ADUsers" {
    BeforeAll {
        # Load script functions if not already loaded (assuming they are globally scoped or loaded)
        . $ScriptPath
    }

    BeforeEach {
        # Set default script-scoped variables used by the function
        $script:TemplateUser = "testtemplate"
        $script:ExportTemplateOnly = $false
        $script:ExportPath = "C:\temp\export.csv"

        # Mock dependencies
        Mock Write-CustomLog { }
        Mock Write-Verbose { }
        Mock Get-ValidADUser { }
        Mock Get-TemplateUserOUs { }
        Mock Get-ADUser { }
        Mock Compare-GroupMembership { return $true }
        Mock Export-Csv { }
    }

    It "Should return early if Get-ValidADUser returns null (Invalid Template User)" {
        Mock Get-ValidADUser { return $null }

        Export-ADUsers

        Assert-MockCalled Get-ValidADUser -Times 1 -Exactly
        Assert-MockCalled Get-TemplateUserOUs -Times 0 -Exactly
    }

    It "Should throw an error if Get-TemplateUserOUs returns null or empty" {
        $mockTemplate = [PSCustomObject]@{ MemberOf = @("Group1") }
        Mock Get-ValidADUser { return $mockTemplate }
        Mock Get-TemplateUserOUs { return $null }

        { Export-ADUsers } | Should -Throw "Keine OUs für Template-Benutzer gefunden"
    }

    It "Should export only the template user if `$ExportTemplateOnly is true" {
        $mockTemplate = [PSCustomObject]@{
            SamAccountName = "testtemplate"
            EmailAddress = "test@example.com"
            MemberOf = @("Group1")
            DistinguishedName = "CN=Test Template,OU=Users,DC=domain,DC=com"
        }
        Mock Get-ValidADUser { return $mockTemplate }
        Mock Get-TemplateUserOUs { return @("OU=Users,DC=domain,DC=com") }

        $script:ExportTemplateOnly = $true

        Export-ADUsers

        Assert-MockCalled Get-ADUser -Times 0 -Exactly
        Assert-MockCalled Export-Csv -Times 1 -Exactly
    }

    It "Should find users, filter by group membership, and export them" {
        $mockTemplate = [PSCustomObject]@{
            SamAccountName = "testtemplate"
            MemberOf = @("Group1")
            DistinguishedName = "CN=Test Template,OU=Users,DC=domain,DC=com"
        }
        Mock Get-ValidADUser { return $mockTemplate }
        Mock Get-TemplateUserOUs { return @("OU=Users,DC=domain,DC=com") }

        $mockUser = [PSCustomObject]@{
            SamAccountName = "founduser"
            EmailAddress = "found@example.com"
            MemberOf = @("Group1")
            DistinguishedName = "CN=Found User,OU=Users,DC=domain,DC=com"
        }
        Mock Get-ADUser { return @($mockUser) } -ParameterFilter { $Filter -eq '*' }
        Mock Compare-GroupMembership { return $true }

        $script:ExportTemplateOnly = $false

        Export-ADUsers

        Assert-MockCalled Get-ADUser -Times 1 -Exactly
        Assert-MockCalled Compare-GroupMembership -Times 1 -Exactly
        Assert-MockCalled Export-Csv -Times 1 -Exactly
    }

    It "Should log a warning and return early if no matching users are found" {
        $mockTemplate = [PSCustomObject]@{
            SamAccountName = "testtemplate"
            MemberOf = @("Group1")
            DistinguishedName = "CN=Test Template,OU=Users,DC=domain,DC=com"
        }
        Mock Get-ValidADUser { return $mockTemplate }
        Mock Get-TemplateUserOUs { return @("OU=Users,DC=domain,DC=com") }

        $mockUser = [PSCustomObject]@{
            SamAccountName = "founduser"
            MemberOf = @("Group2")
            DistinguishedName = "CN=Found User,OU=Users,DC=domain,DC=com"
        }
        Mock Get-ADUser { return @($mockUser) } -ParameterFilter { $Filter -eq '*' }
        Mock Compare-GroupMembership { return $false }

        Export-ADUsers

        Assert-MockCalled Get-ADUser -Times 1 -Exactly
        Assert-MockCalled Export-Csv -Times 0 -Exactly
        Assert-MockCalled Write-CustomLog -ParameterFilter { $Message -eq "Keine Benutzer mit übereinstimmenden Gruppen gefunden" -and $Level -eq "WARNUNG" } -Times 1 -Exactly
    }

    It "Should log a warning and continue if an error occurs during OU search" {
        $mockTemplate = [PSCustomObject]@{
            SamAccountName = "testtemplate"
            MemberOf = @("Group1")
            DistinguishedName = "CN=Test Template,OU=Users,DC=domain,DC=com"
        }
        Mock Get-ValidADUser { return $mockTemplate }
        Mock Get-TemplateUserOUs { return @("OU=Fail,DC=domain,DC=com", "OU=Success,DC=domain,DC=com") }

        # Throw error for first OU, return success for second
        Mock Get-ADUser { throw "Search Error" } -ParameterFilter { $SearchBase -eq "OU=Fail,DC=domain,DC=com" }
        $mockUser = [PSCustomObject]@{
            SamAccountName = "founduser"
            MemberOf = @("Group1")
            DistinguishedName = "CN=Found User,OU=Success,DC=domain,DC=com"
        }
        Mock Get-ADUser { return @($mockUser) } -ParameterFilter { $SearchBase -eq "OU=Success,DC=domain,DC=com" }
        Mock Compare-GroupMembership { return $true }

        Export-ADUsers

        Assert-MockCalled Write-CustomLog -ParameterFilter { $Message -like "Fehler beim Durchsuchen von OU 'OU=Fail,DC=domain,DC=com': Search Error" -and $Level -eq "WARNUNG" } -Times 1 -Exactly
        Assert-MockCalled Export-Csv -Times 1 -Exactly
    }

    It "Should log an error if an exception occurs during export" {
        $mockTemplate = [PSCustomObject]@{
            SamAccountName = "testtemplate"
            MemberOf = @("Group1")
            DistinguishedName = "CN=Test Template,OU=Users,DC=domain,DC=com"
        }
        Mock Get-ValidADUser { return $mockTemplate }
        Mock Get-TemplateUserOUs { return @("OU=Users,DC=domain,DC=com") }

        $mockUser = [PSCustomObject]@{
            SamAccountName = "founduser"
            MemberOf = @("Group1")
            DistinguishedName = "CN=Found User,OU=Users,DC=domain,DC=com"
        }
        Mock Get-ADUser { return @($mockUser) } -ParameterFilter { $Filter -eq '*' }
        Mock Compare-GroupMembership { return $true }

        Mock Export-Csv { throw "Export Error" }

        Export-ADUsers

        Assert-MockCalled Write-CustomLog -ParameterFilter { $Message -like "Fehler beim Exportieren der Benutzerdaten: Export Error" -and $Level -eq "FEHLER" } -Times 1 -Exactly
    }
}
