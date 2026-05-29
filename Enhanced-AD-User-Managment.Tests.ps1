BeforeAll {
    if (-not ("Microsoft.ActiveDirectory.Management.ADUser" -as [type])) {
        Add-Type -TypeDefinition @"
namespace Microsoft.ActiveDirectory.Management {
    public class ADUser {
        public string SamAccountName { get; set; }
        public string Description { get; set; }
        public string Office { get; set; }
        public string Department { get; set; }
        public string Company { get; set; }
        public string EmailAddress { get; set; }
        public string StreetAddress { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string PostalCode { get; set; }
        public string Country { get; set; }
        public string Title { get; set; }
        public string OfficePhone { get; set; }
    }
}
"@
    }

    # Only mock script-specific or missing module functions globally to avoid errors during dot-sourcing
    function global:Write-Log { param($Level, $Message) }
    function global:Add-UserReportEntry { param($SamAccountName, $Status, $Detail) }

    function global:Set-ADUser {
        param(
            $Identity, $Description, $Office, $StreetAddress, $City, $State,
            $PostalCode, $Country, $Department, $Company, $Title,
            $OfficePhone, $EmailAddress
        )
    }
    function global:Get-ADPrincipalGroupMembership { param($Identity) return @() }
    function global:Add-ADPrincipalGroupMembership { param($Identity, $MemberOf) }

    $scriptPath = Join-Path $PSScriptRoot "Enhanced-AD-User-Managment.ps1"

    $ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$null, [ref]$null)
    $functionAst = $ast.Find({
        param($astNode)
        $astNode -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $astNode.Name -eq 'Apply-ADUserProperties'
    }, $true)

    if ($functionAst) {
        $funcBody = $functionAst.Extent.Text
        $sb = [scriptblock]::Create($funcBody)
        . $sb
    } else {
        throw "Could not extract Apply-ADUserProperties function from script"
    }
}

Describe "Apply-ADUserProperties" {
    BeforeEach {
        # Mock Write-Verbose locally if needed by just hiding output
        Mock Write-Verbose { }
    }

    Context "When applying user properties" {
        It "Should call Set-ADUser when properties differ" {
            $refUserObj = New-Object Microsoft.ActiveDirectory.Management.ADUser
            $refUserObj.SamAccountName = "RefUser"
            $refUserObj.Description = "New Description"
            $refUserObj.Office = "New Office"

            $targetUserObj = New-Object Microsoft.ActiveDirectory.Management.ADUser
            $targetUserObj.SamAccountName = "TargetUser"
            $targetUserObj.Description = "Old Description"
            $targetUserObj.Office = "Old Office"

            Mock Set-ADUser { }
            Mock Get-ADPrincipalGroupMembership { @() }

            Apply-ADUserProperties -ReferenceUser $refUserObj -TargetUser $targetUserObj -Confirm:$false

            Assert-MockCalled Set-ADUser -Times 1 -ParameterFilter {
                $Identity.SamAccountName -eq 'TargetUser' -and
                $Description -eq 'New Description' -and
                $Office -eq 'New Office'
            }
        }

        It "Should not call Set-ADUser when properties are identical or empty in ReferenceUser" {
            $refUserObj = New-Object Microsoft.ActiveDirectory.Management.ADUser
            $refUserObj.SamAccountName = "RefUser"
            $refUserObj.Description = "Same Description"
            $refUserObj.Office = ""

            $targetUserObj = New-Object Microsoft.ActiveDirectory.Management.ADUser
            $targetUserObj.SamAccountName = "TargetUser"
            $targetUserObj.Description = "Same Description"
            $targetUserObj.Office = "Existing Office"

            Mock Set-ADUser { }
            Mock Get-ADPrincipalGroupMembership { @() }

            Apply-ADUserProperties -ReferenceUser $refUserObj -TargetUser $targetUserObj -Confirm:$false

            Assert-MockCalled Set-ADUser -Times 0
        }
    }

    Context "When applying group memberships" {
        It "Should add groups that ReferenceUser has but TargetUser is missing" {
            $refUserObj = New-Object Microsoft.ActiveDirectory.Management.ADUser
            $refUserObj.SamAccountName = "RefUser"
            $targetUserObj = New-Object Microsoft.ActiveDirectory.Management.ADUser
            $targetUserObj.SamAccountName = "TargetUser"

            $group1 = [PSCustomObject]@{ Name = "Group1"; DistinguishedName = "CN=Group1" }
            $group2 = [PSCustomObject]@{ Name = "Group2"; DistinguishedName = "CN=Group2" }
            $domainUsers = [PSCustomObject]@{ Name = "Domain Users"; DistinguishedName = "CN=Domain Users" }

            Mock Get-ADPrincipalGroupMembership {
                if ($Identity.SamAccountName -eq "RefUser") {
                    return @($group1, $group2, $domainUsers)
                } else {
                    return @($group1, $domainUsers)
                }
            }
            Mock Set-ADUser { }
            Mock Add-ADPrincipalGroupMembership { }

            Apply-ADUserProperties -ReferenceUser $refUserObj -TargetUser $targetUserObj -Confirm:$false

            Assert-MockCalled Add-ADPrincipalGroupMembership -Times 1 -ParameterFilter {
                $Identity.SamAccountName -eq "TargetUser" -and
                $MemberOf.Name -eq "Group2"
            }
        }
    }

    Context "When error occurs" {
        It "Should catch exception and log error when Set-ADUser fails" {
            $refUserObj = New-Object Microsoft.ActiveDirectory.Management.ADUser
            $refUserObj.SamAccountName = "RefUser"
            $refUserObj.Description = "New Description"

            $targetUserObj = New-Object Microsoft.ActiveDirectory.Management.ADUser
            $targetUserObj.SamAccountName = "TargetUser"
            $targetUserObj.Description = "Old Description"

            Mock Set-ADUser { throw "Simulated AD Error" }
            Mock Get-ADPrincipalGroupMembership { @() }
            Mock Write-Log { }

            Apply-ADUserProperties -ReferenceUser $refUserObj -TargetUser $targetUserObj -Confirm:$false

            Assert-MockCalled Write-Log -ParameterFilter {
                $Level -eq 'Error' -and $Message -match "Simulated AD Error"
            }
        }
    }
}
