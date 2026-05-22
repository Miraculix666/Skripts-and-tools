# Requires -Modules Pester

Describe "New-ADUserFromData" {
    BeforeAll {
        $sutPath = "$PSScriptRoot/../Enhanced-AD-User-Managment.ps1"
        $scriptContent = Get-Content -Path $sutPath -Raw

        $pattern = "(?sm)^[ \t]*function New-ADUserFromData\s*\{.*?\n[ \t]*\}(?=\s*\n\s*# Funktion zum Anwenden)"
        if ($scriptContent -match $pattern) {
            $functionText = $matches[0]
            # Replace AD types with psobject to avoid missing type errors
            $functionText = $functionText -replace '\[Microsoft\.ActiveDirectory\.Management\.ADUser\]', '[psobject]'
            # Replace $PSCmdlet.ShouldProcess with $true for testing
            $functionText = $functionText -replace '\$PSCmdlet\.ShouldProcess\([^\)]+\)', '$true'
            Invoke-Expression $functionText
        } else {
            throw "Could not extract function New-ADUserFromData from $sutPath"
        }
    }

    BeforeEach {
        # Mock global functions that might be called
        function Write-Log { param($Level, $Message) }
        function Add-UserReportEntry { param($SamAccountName, $Status, $Detail) }
        function Write-Verbose { param($Message) }
    }

    It "Should return null and log error if SamAccountName is missing" {
        $userData = @{ GivenName = "John"; Surname = "Doe" }
        $result = New-ADUserFromData -UserData $userData
        $result | Should -BeNullOrEmpty
    }

    It "Should return null if GivenName is missing" {
        $userData = @{ SamAccountName = "jdoe"; Surname = "Doe" }
        $result = New-ADUserFromData -UserData $userData
        $result | Should -BeNullOrEmpty
    }

    It "Should return null if Surname is missing" {
        $userData = @{ SamAccountName = "jdoe"; GivenName = "John" }
        $result = New-ADUserFromData -UserData $userData
        $result | Should -BeNullOrEmpty
    }

    It "Should return null if no target OU can be determined" {
        $userData = @{ SamAccountName = "jdoe"; GivenName = "John"; Surname = "Doe" }
        $result = New-ADUserFromData -UserData $userData
        $result | Should -BeNullOrEmpty
    }

    It "Should return null if user already exists" {
        function Get-ADUser { param($Filter, $ErrorAction) return $true }

        $userData = @{ SamAccountName = "jdoe"; GivenName = "John"; Surname = "Doe"; TargetOU = "OU=Test,DC=Domain,DC=local" }
        $result = New-ADUserFromData -UserData $userData
        $result | Should -BeNullOrEmpty
    }

    It "Should return null if TargetOU does not exist" {
        function Get-ADUser { param($Filter, $ErrorAction) return $false }
        function Get-ADOrganizationalUnit { param($Filter, $ErrorAction) throw "Does not exist" }

        $userData = @{ SamAccountName = "jdoe"; GivenName = "John"; Surname = "Doe"; TargetOU = "OU=Test,DC=Domain,DC=local" }
        $result = New-ADUserFromData -UserData $userData
        $result | Should -BeNullOrEmpty
    }

    It "Should return newly created user object on success" {
        function Get-ADUser { param($Filter, $ErrorAction) return $false }
        function Get-ADOrganizationalUnit { param($Filter, $ErrorAction) return $true }
        function New-ADUser {
            param($SamAccountName, $Name, $GivenName, $Surname, $DisplayName, $UserPrincipalName, $Path, $AccountPassword, $ChangePasswordAtLogon, $Enabled, [switch]$PassThru, $ErrorAction)
            $mockUser = New-Object psobject
            $mockUser | Add-Member -MemberType NoteProperty -Name SamAccountName -Value $SamAccountName
            $mockUser | Add-Member -MemberType NoteProperty -Name SID -Value (New-Object psobject | Add-Member -MemberType NoteProperty -Name Value -Value "S-1-5-21-123" -PassThru)
            return $mockUser
        }

        $userData = @{ SamAccountName = "jdoe"; GivenName = "John"; Surname = "Doe"; TargetOU = "OU=Test,DC=Domain,DC=local" }
        $result = New-ADUserFromData -UserData $userData

        $result | Should -Not -BeNullOrEmpty
        $result.SamAccountName | Should -Be "jdoe"
    }

    It "Should use GlobalTargetOU when provided" {
        function Get-ADUser { param($Filter, $ErrorAction) return $false }
        function Get-ADOrganizationalUnit { param($Filter, $ErrorAction) return $true }
        $script:capturedOU = $null
        function New-ADUser {
            param($SamAccountName, $Name, $GivenName, $Surname, $DisplayName, $UserPrincipalName, $Path, $AccountPassword, $ChangePasswordAtLogon, $Enabled, [switch]$PassThru, $ErrorAction)
            $script:capturedOU = $Path
            $mockUser = New-Object psobject
            $mockUser | Add-Member -MemberType NoteProperty -Name SamAccountName -Value $SamAccountName
            $mockUser | Add-Member -MemberType NoteProperty -Name SID -Value (New-Object psobject | Add-Member -MemberType NoteProperty -Name Value -Value "S-1-5-21-123" -PassThru)
            return $mockUser
        }

        $userData = @{ SamAccountName = "jdoe"; GivenName = "John"; Surname = "Doe" }
        $result = New-ADUserFromData -UserData $userData -GlobalTargetOU "OU=Global,DC=Domain,DC=local"

        $result | Should -Not -BeNullOrEmpty
        $script:capturedOU | Should -Be "OU=Global,DC=Domain,DC=local"
    }

    It "Should derive target OU from TemplateUser" {
        function Get-ADUser { param($Filter, $ErrorAction) return $false }
        function Get-ADOrganizationalUnit { param($Filter, $ErrorAction) return $true }
        $script:capturedOU = $null
        function New-ADUser {
            param($SamAccountName, $Name, $GivenName, $Surname, $DisplayName, $UserPrincipalName, $Path, $AccountPassword, $ChangePasswordAtLogon, $Enabled, [switch]$PassThru, $ErrorAction)
            $script:capturedOU = $Path
            $mockUser = New-Object psobject
            $mockUser | Add-Member -MemberType NoteProperty -Name SamAccountName -Value $SamAccountName
            $mockUser | Add-Member -MemberType NoteProperty -Name SID -Value (New-Object psobject | Add-Member -MemberType NoteProperty -Name Value -Value "S-1-5-21-123" -PassThru)
            return $mockUser
        }

        $templateUser = New-Object psobject
        $templateUser | Add-Member -MemberType NoteProperty -Name DistinguishedName -Value "CN=Template User,OU=TemplateOU,DC=Domain,DC=local"

        $userData = @{ SamAccountName = "jdoe"; GivenName = "John"; Surname = "Doe" }
        $result = New-ADUserFromData -UserData $userData -TemplateUser $templateUser

        $result | Should -Not -BeNullOrEmpty
        $script:capturedOU | Should -Be "OU=TemplateOU,DC=Domain,DC=local"
    }

    It "Should apply standard attributes and attributes from UserData" {
        function Get-ADUser { param($Filter, $ErrorAction) return $false }
        function Get-ADOrganizationalUnit { param($Filter, $ErrorAction) return $true }
        $script:capturedParams = $null
        function New-ADUser {
            param($SamAccountName, $Name, $GivenName, $Surname, $DisplayName, $UserPrincipalName, $Path, $AccountPassword, $ChangePasswordAtLogon, $Enabled, [switch]$PassThru, $ErrorAction, $Title, $Department)
            $script:capturedParams = $MyInvocation.BoundParameters
            $mockUser = New-Object psobject
            $mockUser | Add-Member -MemberType NoteProperty -Name SamAccountName -Value $SamAccountName
            $mockUser | Add-Member -MemberType NoteProperty -Name SID -Value (New-Object psobject | Add-Member -MemberType NoteProperty -Name Value -Value "S-1-5-21-123" -PassThru)
            return $mockUser
        }

        $userData = @{
            SamAccountName = "jdoe"
            GivenName = "John"
            Surname = "Doe"
            TargetOU = "OU=Test,DC=Domain,DC=local"
            Title = "Manager"
            Department = "IT"
        }
        $result = New-ADUserFromData -UserData $userData

        $result | Should -Not -BeNullOrEmpty
        $script:capturedParams["Title"] | Should -Be "Manager"
        $script:capturedParams["Department"] | Should -Be "IT"
        $script:capturedParams["Name"] | Should -Be "John Doe"
        $script:capturedParams["DisplayName"] | Should -Be "John Doe"
    }

    It "Should assign groups from TemplateUser if successfully created" {
        function Get-ADUser { param($Filter, $ErrorAction) return $false }
        function Get-ADOrganizationalUnit { param($Filter, $ErrorAction) return $true }
        function New-ADUser {
            param($SamAccountName, $Name, $GivenName, $Surname, $DisplayName, $UserPrincipalName, $Path, $AccountPassword, $ChangePasswordAtLogon, $Enabled, [switch]$PassThru, $ErrorAction)
            $mockUser = New-Object psobject
            $mockUser | Add-Member -MemberType NoteProperty -Name SamAccountName -Value $SamAccountName
            $mockUser | Add-Member -MemberType NoteProperty -Name SID -Value (New-Object psobject | Add-Member -MemberType NoteProperty -Name Value -Value "S-1-5-21-123" -PassThru)
            return $mockUser
        }
        function Get-ADPrincipalGroupMembership {
            param($Identity, $ErrorAction)
            $group1 = New-Object psobject | Add-Member -MemberType NoteProperty -Name Name -Value "IT-Admins" -PassThru
            $group2 = New-Object psobject | Add-Member -MemberType NoteProperty -Name Name -Value "Domain Users" -PassThru
            return @($group1, $group2)
        }
        $script:addedGroups = $null
        function Add-ADPrincipalGroupMembership {
            param($Identity, $MemberOf, $ErrorAction)
            $script:addedGroups = $MemberOf
        }

        $templateUser = New-Object psobject
        $templateUser | Add-Member -MemberType NoteProperty -Name DistinguishedName -Value "CN=Template User,OU=TemplateOU,DC=Domain,DC=local"
        $templateUser | Add-Member -MemberType NoteProperty -Name SamAccountName -Value "template_user"

        $userData = @{ SamAccountName = "jdoe"; GivenName = "John"; Surname = "Doe" }
        $result = New-ADUserFromData -UserData $userData -TemplateUser $templateUser

        $result | Should -Not -BeNullOrEmpty
        $script:addedGroups.Count | Should -Be 1
        $script:addedGroups[0].Name | Should -Be "IT-Admins"
    }
}
