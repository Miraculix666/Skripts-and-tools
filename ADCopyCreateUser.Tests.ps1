$ErrorActionPreference = 'Stop'

# Define dummy ADUser class to allow static analysis/loading if the module isn't present
if (-not ("Microsoft.ActiveDirectory.Management.ADUser" -as [type])) {
    Add-Type -TypeDefinition @"
    namespace Microsoft.ActiveDirectory.Management {
        public class ADUser {
            public string SamAccountName;
            public string DistinguishedName;
        }
    }
"@ -IgnoreWarnings
}

# Dynamically extract functions from the script without executing it
$scriptPath = Join-Path $PSScriptRoot 'ADCopyCreateUser.ps1'

# We use the AST to find and extract the function definitions so we test the live code
# without executing the script's outer Process block.
$ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$null, [ref]$null)
$functionAsts = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)

foreach ($funcAst in $functionAsts) {
    if ($funcAst.Name -in @('Copy-ADUserWithGroups', 'Write-Log')) {
        # Create the function in the current scope
        Invoke-Expression "function $($funcAst.Name) { $($funcAst.Body.Extent.Text) }"
    }
}

Describe "Copy-ADUserWithGroups" {
    BeforeAll {
        # Dummy Source User
        $global:sourceUser = New-Object Microsoft.ActiveDirectory.Management.ADUser
        $global:sourceUser.SamAccountName = "source.user"
        $global:sourceUser.DistinguishedName = "CN=Source User,OU=SourceOU,DC=domain,DC=com"

        $global:targetParams = @{
            SamAccountName = "target.user"
            Enabled = $true
        }

        # Dummy Target User (returned by New-ADUser)
        $global:targetUser = New-Object Microsoft.ActiveDirectory.Management.ADUser
        $global:targetUser.SamAccountName = "target.user"
        $global:targetUser.DistinguishedName = "CN=Target User,OU=DefaultUsersOU,DC=domain,DC=com"
    }

    Context "Happy Path - User with groups" {
        It "Should create user, copy groups, and move OU correctly" {
            Mock Write-Log {}

            Mock New-ADUser { return $global:targetUser }

            Mock Get-ADPrincipalGroupMembership {
                return @(
                    [pscustomobject]@{ ObjectClass = 'group'; Name = 'Group1' },
                    [pscustomobject]@{ ObjectClass = 'group'; Name = 'Group2' },
                    [pscustomobject]@{ ObjectClass = 'user'; Name = 'NotAGroup' } # Should be filtered out
                )
            }

            Mock Add-ADPrincipalGroupMembership {}

            Mock Move-ADObject {}

            $result = Copy-ADUserWithGroups -SourceUser $global:sourceUser -TargetUserParams $global:targetParams

            $result.SamAccountName | Should -Be "target.user"

            # Assert New-ADUser was called correctly
            Assert-MockCalled New-ADUser -Times 1 -ParameterFilter {
                $Instance -eq $global:sourceUser -and $SamAccountName -eq 'target.user'
            }

            # Assert Get-ADPrincipalGroupMembership was called
            Assert-MockCalled Get-ADPrincipalGroupMembership -Times 1 -ParameterFilter {
                $Identity -eq $global:sourceUser
            }

            # Assert Add-ADPrincipalGroupMembership was called with only groups
            Assert-MockCalled Add-ADPrincipalGroupMembership -Times 1 -ParameterFilter {
                $Identity -eq $global:targetUser -and $MemberOf.Count -eq 2
            }

            # Assert Move-ADObject was called with correctly extracted OU path
            Assert-MockCalled Move-ADObject -Times 1 -ParameterFilter {
                $Identity -eq $global:targetUser.DistinguishedName -and $TargetPath -eq "OU=SourceOU,DC=domain,DC=com"
            }
        }
    }

    Context "Edge Case - Source user has no groups" {
        It "Should create user and move OU, but not attempt to copy groups" {
            Mock Write-Log {}
            Mock New-ADUser { return $global:targetUser }
            Mock Get-ADPrincipalGroupMembership { return @() }
            Mock Add-ADPrincipalGroupMembership {}
            Mock Move-ADObject {}

            $result = Copy-ADUserWithGroups -SourceUser $global:sourceUser -TargetUserParams $global:targetParams

            $result.SamAccountName | Should -Be "target.user"

            Assert-MockCalled New-ADUser -Times 1
            Assert-MockCalled Get-ADPrincipalGroupMembership -Times 1
            Assert-MockCalled Add-ADPrincipalGroupMembership -Times 0
            Assert-MockCalled Move-ADObject -Times 1
        }
    }

    Context "Error Handling" {
        It "Should catch exception from New-ADUser, log it, and throw" {
            Mock Write-Log {}
            Mock New-ADUser { throw "Simulated AD Error" }

            { Copy-ADUserWithGroups -SourceUser $global:sourceUser -TargetUserParams $global:targetParams } | Should -Throw "Simulated AD Error"

            Assert-MockCalled Write-Log -ParameterFilter { $Level -eq 'Error' -and $Message -match "Simulated AD Error" }
        }
    }
}
