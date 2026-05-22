$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptPath = Join-Path $here "UserCopyCreate.ps1"

Describe "New-CustomADUser" {
    BeforeAll {
        # Mock functions that exist in the script but not in our test environment
        function Write-CustomLog { param($Message, $Level) }

        # Load the script to test
        . $scriptPath
    }

    Context "Happy Path" {
        It "creates a user with valid parameters when OU exists" {
            # Arrange
            Mock Get-ADOrganizationalUnit { return @{ DistinguishedName = "OU=Test,DC=Domain,DC=com" } }
            Mock New-ADUser {}
            Mock Write-Verbose {}

            $pass = ConvertTo-SecureString "TestPass123!" -AsPlainText -Force

            # Act
            New-CustomADUser -SamAccountName "jdoe" -UserPrincipalName "jdoe@domain.com" -Name "John Doe" -OU "OU=Test,DC=Domain,DC=com" -Password $pass -GivenName "John"

            # Assert
            Assert-MockCalled New-ADUser -Times 1 -ParameterFilter {
                $SamAccountName -eq "jdoe" -and
                $Name -eq "John Doe" -and
                $GivenName -eq "John"
            }
        }

        It "adds user to multiple groups if provided" {
            # Arrange
            Mock Get-ADOrganizationalUnit { return @{ DistinguishedName = "OU=Test,DC=Domain,DC=com" } }
            Mock New-ADUser {}
            Mock Add-ADGroupMember {}
            Mock Write-Verbose {}

            $pass = ConvertTo-SecureString "TestPass123!" -AsPlainText -Force
            $groups = @("Group1", "Group2")

            # Act
            New-CustomADUser -SamAccountName "jdoe" -UserPrincipalName "jdoe@domain.com" -Name "John Doe" -OU "OU=Test,DC=Domain,DC=com" -Password $pass -Groups $groups

            # Assert
            Assert-MockCalled Add-ADGroupMember -Times 2
            Assert-MockCalled Add-ADGroupMember -Times 1 -ParameterFilter { $Identity -eq "Group1" -and $Members -eq "jdoe" }
            Assert-MockCalled Add-ADGroupMember -Times 1 -ParameterFilter { $Identity -eq "Group2" -and $Members -eq "jdoe" }
        }
    }

    Context "Error Handling" {
        It "throws an error if OU does not exist" {
            # Arrange
            Mock Get-ADOrganizationalUnit { return $null }
            Mock Write-CustomLog {}
            Mock Write-Verbose {}

            $pass = ConvertTo-SecureString "TestPass123!" -AsPlainText -Force

            # Act & Assert
            { New-CustomADUser -SamAccountName "jdoe" -UserPrincipalName "jdoe@domain.com" -Name "John Doe" -OU "OU=Invalid,DC=Domain,DC=com" -Password $pass } | Should -Throw "Die angegebene OU existiert nicht: OU=Invalid,DC=Domain,DC=com"

            # Write-CustomLog should be called in the catch block
            Assert-MockCalled Write-CustomLog -Times 1 -ParameterFilter { $Level -eq "FEHLER" }
        }

        It "catches and logs errors during group assignment without throwing" {
            # Arrange
            Mock Get-ADOrganizationalUnit { return @{ DistinguishedName = "OU=Test,DC=Domain,DC=com" } }
            Mock New-ADUser {}
            Mock Add-ADGroupMember { throw "Group does not exist" }
            Mock Write-CustomLog {}
            Mock Write-Verbose {}

            $pass = ConvertTo-SecureString "TestPass123!" -AsPlainText -Force
            $groups = @("InvalidGroup")

            # Act (Should not throw since try-catch is inside the loop)
            New-CustomADUser -SamAccountName "jdoe" -UserPrincipalName "jdoe@domain.com" -Name "John Doe" -OU "OU=Test,DC=Domain,DC=com" -Password $pass -Groups $groups

            # Assert
            Assert-MockCalled Write-CustomLog -Times 1 -ParameterFilter { $Level -eq "WARNUNG" -and $Message -match "Fehler beim Hinzufügen von Benutzer" }
        }

        It "throws an error if New-ADUser fails" {
            # Arrange
            Mock Get-ADOrganizationalUnit { return @{ DistinguishedName = "OU=Test,DC=Domain,DC=com" } }
            Mock New-ADUser { throw "AD User already exists" }
            Mock Write-CustomLog {}
            Mock Write-Verbose {}

            $pass = ConvertTo-SecureString "TestPass123!" -AsPlainText -Force

            # Act & Assert
            { New-CustomADUser -SamAccountName "jdoe" -UserPrincipalName "jdoe@domain.com" -Name "John Doe" -OU "OU=Test,DC=Domain,DC=com" -Password $pass } | Should -Throw "AD User already exists"

            # Ensure it logs the failure
            Assert-MockCalled Write-CustomLog -Times 1 -ParameterFilter { $Level -eq "FEHLER" -and $Message -match "Fehler bei der Erstellung des Benutzers" }
        }
    }
}
