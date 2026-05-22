BeforeAll {
    # Mocking AD functions and Write-Log(Message) functions to avoid side effects
    function Get-ADUser {}
    function ConvertTo-SecureString {}
    function New-ADUser {}
    function Write-Log {}
    function Write-LogMessage {}

    . "$PSScriptRoot/../CopyCreateNewUser.ps1"
}

Describe "Process-UserCreation" {
    Context "Successful User Creation" {
        It "should successfully create a new user with standard parameters" {
            Mock Get-ADUser { return @{ DistinguishedName = "CN=Template,OU=Templates,DC=domain,DC=com" } }
            Mock ConvertTo-SecureString { return "SecureStringMock" }
            Mock New-ADUser {}
            Mock Write-Log {}
            Mock Write-LogMessage {}

            $env:USERDNSDOMAIN = "domain.com"

            $userData = @{
                SamAccountName = "testuser"
                GivenName = "Test"
                Surname = "User"
                Password = "Password123"
            }
            $templateUser = "templateuser"
            $targetOU = "OU=Users,DC=domain,DC=com"

            Process-UserCreation -UserData $userData -TemplateUser $templateUser -TargetOU $targetOU

            Assert-MockCalled New-ADUser -Times 1 -Exactly
            Assert-MockCalled Write-LogMessage -Times 1 -ParameterFilter { $Message -like "*erfolgreich erstellt*" }
        }

        It "should include optional parameters if they exist in UserData" {
            Mock Get-ADUser { return @{ DistinguishedName = "CN=Template,OU=Templates,DC=domain,DC=com" } }
            Mock ConvertTo-SecureString { return "SecureStringMock" }
            Mock New-ADUser {}
            Mock Write-Log {}
            Mock Write-LogMessage {}

            $env:USERDNSDOMAIN = "domain.com"

            $userData = @{
                SamAccountName = "testuser"
                GivenName = "Test"
                Surname = "User"
                Password = "Password123"
                Department = "IT"
                Title = "Engineer"
            }
            $templateUser = "templateuser"
            $targetOU = "OU=Users,DC=domain,DC=com"

            Process-UserCreation -UserData $userData -TemplateUser $templateUser -TargetOU $targetOU

            Assert-MockCalled New-ADUser -Times 1 -Exactly -ParameterFilter { $Department -eq "IT" -and $Title -eq "Engineer" }
        }
    }

    Context "User Creation Failure" {
        It "should handle errors and log them when New-ADUser fails" {
            Mock Get-ADUser { return @{ DistinguishedName = "CN=Template,OU=Templates,DC=domain,DC=com" } }
            Mock ConvertTo-SecureString { return "SecureStringMock" }
            Mock New-ADUser { throw "AD Error" }
            Mock Write-Log {}
            Mock Write-LogMessage {}

            $env:USERDNSDOMAIN = "domain.com"

            $userData = @{
                SamAccountName = "failuser"
                GivenName = "Fail"
                Surname = "User"
                Password = "Password123"
            }
            $templateUser = "templateuser"
            $targetOU = "OU=Users,DC=domain,DC=com"

            Process-UserCreation -UserData $userData -TemplateUser $templateUser -TargetOU $targetOU

            Assert-MockCalled New-ADUser -Times 1 -Exactly
            Assert-MockCalled Write-LogMessage -Times 1 -ParameterFilter { $Message -like "*Fehler beim Erstellen*" -and $Type -eq "Error" }
        }
    }
}
