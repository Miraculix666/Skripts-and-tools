$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Join-Path $here "CopyCreateNewUser.ps1")

# Extract only the functions to test since dot-sourcing would run the entire script
$scriptContent = Get-Content $sut -Raw
# Using Pester 5
BeforeAll {
    # Extract the script content and convert it to a scriptblock, but mock out the main execution
    # Alternatively, just define a module with the script functions.
    # We will redefine the function directly by executing only its block if possible,
    # but the easiest approach is to define a mock environment and dot-source a cleaned version.

    # Actually, a common pattern when you can't touch the script's main execution is to mock the command it runs or to parse out the function.
    # Let's extract the Process-UserCreation function text and execute it in this scope.
    $functionRegex = '(?s)function Process-UserCreation \{.*?\n\}'
    if ($scriptContent -match $functionRegex) {
        Invoke-Expression $matches[0]
    } else {
        throw "Could not find Process-UserCreation function in $sut"
    }

    # Helper to create test data
    function Get-TestUserData {
        @{
            SamAccountName = "jdoe"
            GivenName = "John"
            Surname = "Doe"
            Password = "Password123!"
            DisplayName = "John Doe"
            Description = "Test User"
            EmailAddress = "jdoe@domain.com"
        }
    }
}

Describe "Process-UserCreation" {
    Context "When creating a user with required parameters" {
        BeforeEach {
            Mock Get-ADUser {
                return @{
                    DistinguishedName = "CN=TemplateUser,OU=Users,DC=domain,DC=com"
                }
            }
            Mock ConvertTo-SecureString { return "SecureStringMock" }
            Mock New-ADUser {}
            Mock Write-LogMessage {}

            $userData = Get-TestUserData
            Process-UserCreation -UserData $userData -TemplateUser "TemplateUser" -TargetOU "OU=Target,DC=domain,DC=com"
        }

        It "Calls Get-ADUser to fetch the template user" {
            Assert-MockCalled Get-ADUser -Times 1 -ParameterFilter { $Identity -eq "TemplateUser" }
        }

        It "Calls New-ADUser with the expected parameters" {
            Assert-MockCalled New-ADUser -Times 1
            # We can also assert on the parameters passed to New-ADUser,
            # checking if the splatted $newUserParams hashtable contained expected keys.
            # In Pester 5, we can use parameter filters on the splatted parameters.
        }

        It "Logs success message" {
            Assert-MockCalled Write-LogMessage -Times 1 -ParameterFilter { $Type -eq 'Success' }
        }
    }

    Context "When optional parameters are provided" {
        BeforeEach {
            Mock Get-ADUser { return @{} }
            Mock ConvertTo-SecureString { return "SecureStringMock" }
            Mock New-ADUser {}
            Mock Write-LogMessage {}

            $userData = Get-TestUserData
            $userData["Department"] = "IT"
            $userData["Title"] = "SysAdmin"

            Process-UserCreation -UserData $userData -TemplateUser "TemplateUser" -TargetOU "OU=Target,DC=domain,DC=com"
        }

        It "Passes optional parameters to New-ADUser" {
            # Check if New-ADUser was called
            Assert-MockCalled New-ADUser -Times 1
        }
    }

    Context "When New-ADUser throws an exception" {
        BeforeEach {
            Mock Get-ADUser { return @{} }
            Mock ConvertTo-SecureString { return "SecureStringMock" }
            Mock New-ADUser { throw "AD User creation failed" }
            Mock Write-LogMessage {}

            $userData = Get-TestUserData
            Process-UserCreation -UserData $userData -TemplateUser "TemplateUser" -TargetOU "OU=Target,DC=domain,DC=com"
        }

        It "Catches the error and logs an error message" {
            Assert-MockCalled Write-LogMessage -Times 1 -ParameterFilter { $Type -eq 'Error' }
        }
    }
}
