$scriptPath = Join-Path $PSScriptRoot "UserPWReset.ps1"

Describe "ResetPasswordAndEnableAccount" {
    BeforeAll {
        # Mock global commands and cmdlets to prevent side-effects during dot-sourcing
        Mock Import-Module {}
        Mock Get-CimInstance { return [pscustomobject]@{Domain='TestDomain'} }
        Mock Get-ADUser { return @() }
        Mock Export-Csv {}
        Mock Start-Transcript {}
        Mock Stop-Transcript {}
        Mock Write-Host {}
        Mock Write-Verbose {}
        Mock Read-Host { return 'test' }

        # Dot-source the script to load the function, avoiding interactive prompts
        . $scriptPath -UserName "testuser" -Password "testpass" -SkipConfirmation
    }

    Context "Happy Path - Successfully resets password and enables account" {
        BeforeEach {
            Mock ConvertTo-SecureString { return 'SecureStringMock' }
            Mock Set-ADAccountPassword {}
            Mock Enable-ADAccount {}
            Mock Set-ADAccountControl {}
            Mock Write-Host {}
            Mock Write-Verbose {}
        }

        It "Should convert the password to a secure string" {
            $Script:Silent = $false
            ResetPasswordAndEnableAccount -Identity "testuser" -NewPassword "newpass"

            Should -Invoke -CommandName ConvertTo-SecureString -Times 1 -ParameterFilter {
                $String -eq 'newpass' -and $AsPlainText -and $Force
            }
        }

        It "Should call AD cmdlets with correct parameters" {
            $Script:Silent = $false
            ResetPasswordAndEnableAccount -Identity "testuser" -NewPassword "newpass"

            Should -Invoke -CommandName Set-ADAccountPassword -Times 1 -ParameterFilter {
                $Identity -eq 'testuser' -and $Reset -and $NewPassword -eq 'SecureStringMock' -and $ErrorAction -eq 'Stop'
            }
            Should -Invoke -CommandName Enable-ADAccount -Times 1 -ParameterFilter {
                $Identity -eq 'testuser' -and $ErrorAction -eq 'Stop'
            }
            Should -Invoke -CommandName Set-ADAccountControl -Times 1 -ParameterFilter {
                $Identity -eq 'testuser' -and $CannotChangePassword -eq $true -and $ErrorAction -eq 'Stop'
            }
        }

        It "Should write a success message to Write-Host when not silent" {
            $Script:Silent = $false
            ResetPasswordAndEnableAccount -Identity "testuser" -NewPassword "newpass"

            Should -Invoke -CommandName Write-Host -Times 1 -ParameterFilter {
                $Object -match 'Password reset and account enabled for testuser' -and $ForegroundColor -eq 'Green'
            }
        }

        It "Should not write to Write-Host when silent" {
            $Script:Silent = $true
            ResetPasswordAndEnableAccount -Identity "testuser" -NewPassword "newpass"

            Should -Invoke -CommandName Write-Host -Times 0 -ParameterFilter {
                $Object -match 'Password reset'
            }
        }

        It "Should write to Write-Verbose regardless of Silent switch" {
            $Script:Silent = $true
            ResetPasswordAndEnableAccount -Identity "testuser" -NewPassword "newpass"

            Should -Invoke -CommandName Write-Verbose -Times 1 -ParameterFilter {
                $Message -match 'Password reset and account enabled for testuser'
            }
        }
    }

    Context "Error Path - Fails to reset password or enable account" {
        BeforeEach {
            Mock ConvertTo-SecureString { return 'SecureStringMock' }
            Mock Set-ADAccountPassword { throw "Simulated Error" }
            Mock Enable-ADAccount {}
            Mock Set-ADAccountControl {}
            Mock Write-Host {}
            Mock Write-Verbose {}
        }

        It "Should catch the error and not proceed with enabling account" {
            $Script:Silent = $false
            ResetPasswordAndEnableAccount -Identity "erroruser" -NewPassword "newpass"

            Should -Invoke -CommandName Set-ADAccountPassword -Times 1 -ParameterFilter { $Identity -eq 'erroruser' }
            Should -Invoke -CommandName Enable-ADAccount -Times 0
            Should -Invoke -CommandName Set-ADAccountControl -Times 0
        }

        It "Should write an error message to Write-Host when not silent" {
            $Script:Silent = $false
            ResetPasswordAndEnableAccount -Identity "erroruser" -NewPassword "newpass"

            Should -Invoke -CommandName Write-Host -Times 1 -ParameterFilter {
                $Object -match 'Failed to reset password and enable account for erroruser' -and $ForegroundColor -eq 'Red'
            }
        }

        It "Should not write an error message to Write-Host when silent" {
            $Script:Silent = $true
            ResetPasswordAndEnableAccount -Identity "erroruser" -NewPassword "newpass"

            Should -Invoke -CommandName Write-Host -Times 0 -ParameterFilter {
                $Object -match 'Failed to reset password'
            }
        }

        It "Should write error to Write-Verbose regardless of Silent switch" {
            $Script:Silent = $true
            ResetPasswordAndEnableAccount -Identity "erroruser" -NewPassword "newpass"

            Should -Invoke -CommandName Write-Verbose -Times 1 -ParameterFilter {
                $Message -match 'Failed to reset password and enable account for erroruser'
            }
        }
    }
}
