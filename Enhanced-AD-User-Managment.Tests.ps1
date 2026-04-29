$ScriptPath = Join-Path $PSScriptRoot "Enhanced-AD-User-Managment.ps1"

# Pester 5 structure
BeforeAll {
    # To dot-source the script without running its process block, we might need to be careful.
    # The script has `begin`, `process`, and `end` blocks and doesn't run the process block unless arguments are passed.
    # However, to be safe, we can mock cmdlets that would be called if it did run,
    # or rely on dot-sourcing just loading functions.
    . $ScriptPath
}

Describe "New-ADUserFromData" {
    BeforeAll {
        # Mock logging and reporting globally for the describe block
        Mock Write-Log {}
        Mock Add-UserReportEntry {}
    }

    Context "Missing Required Data" {
        It "Returns `$null and logs error when SamAccountName is missing" {
            $userData = @{ GivenName = "John"; Surname = "Doe" }
            $result = New-ADUserFromData -UserData $userData
            $result | Should -BeNullOrEmpty
            Assert-MockCalled Write-Log -Times 1 -ParameterFilter { $Level -eq 'Error' -and $Message -like "*SamAccountName*" }
        }

        It "Returns `$null and logs error when GivenName is missing" {
            $userData = @{ SamAccountName = "jdoe"; Surname = "Doe" }
            $result = New-ADUserFromData -UserData $userData
            $result | Should -BeNullOrEmpty
            Assert-MockCalled Write-Log -Times 1 -ParameterFilter { $Level -eq 'Error' -and $Message -like "*GivenName*" }
        }

        It "Returns `$null and logs error when Surname is missing" {
            $userData = @{ SamAccountName = "jdoe"; GivenName = "John" }
            $result = New-ADUserFromData -UserData $userData
            $result | Should -BeNullOrEmpty
            Assert-MockCalled Write-Log -Times 1 -ParameterFilter { $Level -eq 'Error' -and $Message -like "*Surname*" }
        }
    }

    Context "TargetOU Determination" {
        It "Returns `$null and logs error when TargetOU cannot be determined" {
            $userData = @{ SamAccountName = "jdoe"; GivenName = "John"; Surname = "Doe" }
            # No GlobalTargetOU, no TargetOU in CSV, no TemplateUser
            $result = New-ADUserFromData -UserData $userData
            $result | Should -BeNullOrEmpty
            Assert-MockCalled Write-Log -Times 1 -ParameterFilter { $Level -eq 'Error' -and $Message -like "*Keine Ziel-OU*" }
        }
    }

    Context "Password Handling" {
        It "Returns `$null when password conversion fails" {
            $userData = @{ SamAccountName = "jdoe"; GivenName = "John"; Surname = "Doe"; Password = "InvalidPasswordFormatForSomeReason" }

            Mock ConvertTo-SecureString { throw "Conversion Failed" }

            $result = New-ADUserFromData -UserData $userData -GlobalTargetOU "OU=Test,DC=Domain,DC=com"
            $result | Should -BeNullOrEmpty
            Assert-MockCalled Write-Log -Times 1 -ParameterFilter { $Level -eq 'Error' -and $Message -like "*Fehler beim Konvertieren des Passworts*" }
        }
    }

    Context "Active Directory Validations" {
        BeforeEach {
            Mock ConvertTo-SecureString { return "SecureStringMock" }
        }

        It "Returns `$null when target user already exists" {
            $userData = @{ SamAccountName = "jdoe"; GivenName = "John"; Surname = "Doe" }
            Mock Get-ADUser { return [PSCustomObject]@{ SamAccountName = "jdoe" } }

            $result = New-ADUserFromData -UserData $userData -GlobalTargetOU "OU=Test,DC=Domain,DC=com"
            $result | Should -BeNullOrEmpty
            Assert-MockCalled Write-Log -Times 1 -ParameterFilter { $Level -eq 'Error' -and $Message -like "*existiert bereits*" }
        }

        It "Returns `$null when TargetOU does not exist" {
            $userData = @{ SamAccountName = "jdoe"; GivenName = "John"; Surname = "Doe" }
            Mock Get-ADUser { throw "User not found" } # Normal behavior when user doesn't exist
            Mock Get-ADOrganizationalUnit { throw "OU not found" }

            $result = New-ADUserFromData -UserData $userData -GlobalTargetOU "OU=Invalid,DC=Domain,DC=com"
            $result | Should -BeNullOrEmpty
            Assert-MockCalled Write-Log -Times 1 -ParameterFilter { $Level -eq 'Error' -and $Message -like "*Die Ziel-OU*existiert nicht*" }
        }
    }

    Context "Happy Path User Creation" {
        BeforeEach {
            Mock ConvertTo-SecureString { return "SecureStringMock" }
            Mock Get-ADUser { throw "User not found" } # Normal behavior when user doesn't exist
            Mock Get-ADOrganizationalUnit { return [PSCustomObject]@{ DistinguishedName = "OU=Test,DC=Domain,DC=com" } }
            Mock New-ADUser { return [PSCustomObject]@{ SamAccountName = $SamAccountName; SID = [PSCustomObject]@{ Value = "S-1-5-21-MockSID" } } }
        }

        It "Successfully creates user when all required info is provided" {
            $userData = @{ SamAccountName = "jdoe"; GivenName = "John"; Surname = "Doe" }

            $result = New-ADUserFromData -UserData $userData -GlobalTargetOU "OU=Test,DC=Domain,DC=com"

            $result | Should -Not -BeNullOrEmpty
            $result.SamAccountName | Should -Be "jdoe"
            Assert-MockCalled New-ADUser -Times 1
            Assert-MockCalled Write-Log -Times 1 -ParameterFilter { $Level -eq 'Info' -and $Message -like "*erfolgreich erstellt*" }
        }

        It "Successfully copies groups from TemplateUser when provided" {
            $userData = @{ SamAccountName = "jdoe"; GivenName = "John"; Surname = "Doe" }

            $templateUser = [PSCustomObject]@{
                SamAccountName = "template"
                DistinguishedName = "CN=template,OU=Test,DC=Domain,DC=com"
                PSObject = [PSCustomObject]@{
                    Properties = [PSCustomObject]@{
                        Match = { param($name) if ($name -eq "Description") { return @(1) } else { return @() } }
                    }
                }
                Description = "Template Description"
            }
            # Give template user groups
            Mock Get-ADPrincipalGroupMembership {
                return @(
                    [PSCustomObject]@{ Name = "Domain Users" },
                    [PSCustomObject]@{ Name = "Accounting" }
                )
            }
            Mock Add-ADPrincipalGroupMembership {}

            # Using $templateUser overrides GlobalTargetOU check, but falls back to pulling OU from TemplateUser
            $result = New-ADUserFromData -UserData $userData -TemplateUser $templateUser

            $result | Should -Not -BeNullOrEmpty
            Assert-MockCalled Add-ADPrincipalGroupMembership -Times 1
            Assert-MockCalled Write-Log -Times 1 -ParameterFilter { $Level -eq 'Info' -and $Message -like "*Gruppenmitgliedschaften*erfolgreich hinzugefügt*" }
        }
    }
}
