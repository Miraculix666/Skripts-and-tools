BeforeAll {
    $scriptPath = "$PSScriptRoot/ChangeUserProps.ps1"
    $scriptAST = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$null, [ref]$null)
    $functionAST = $scriptAST.Find({$args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $args[0].Name -eq "Test-ADModule"}, $true)
    Invoke-Expression $functionAST.Extent.Text
}

Describe "Test-ADModule" {

    Context "When ActiveDirectory module is not available (not installed)" {
        It "should return false and write an error" {
            Mock Get-Module { return $null } -ParameterFilter { $ListAvailable -and $Name -eq 'ActiveDirectory' }
            Mock Write-Error {} # Suppress error output during test

            $result = Test-ADModule

            $result | Should -Be $false
            Assert-MockCalled Get-Module -Times 1 -ParameterFilter { $ListAvailable -and $Name -eq 'ActiveDirectory' }
            Assert-MockCalled Write-Error -Times 1
        }
    }

    Context "When ActiveDirectory module is already loaded" {
        It "should return true without importing the module" {
            Mock Get-Module { return @{ Name = 'ActiveDirectory' } } -ParameterFilter { $ListAvailable -and $Name -eq 'ActiveDirectory' }
            Mock Get-Module { return @{ Name = 'ActiveDirectory' } } -ParameterFilter { $Name -eq 'ActiveDirectory' -and -not $ListAvailable }
            Mock Import-Module {}

            $result = Test-ADModule

            $result | Should -Be $true
            Assert-MockCalled Get-Module -Times 1 -ParameterFilter { $ListAvailable -and $Name -eq 'ActiveDirectory' }
            Assert-MockCalled Get-Module -Times 1 -ParameterFilter { $Name -eq 'ActiveDirectory' -and -not $ListAvailable }
            Assert-MockCalled Import-Module -Times 0
        }
    }

    Context "When ActiveDirectory module is available but not loaded, and loads successfully" {
        It "should return true and import the module" {
            Mock Get-Module { return @{ Name = 'ActiveDirectory' } } -ParameterFilter { $ListAvailable -and $Name -eq 'ActiveDirectory' }
            Mock Get-Module { return $null } -ParameterFilter { $Name -eq 'ActiveDirectory' -and -not $ListAvailable }
            Mock Import-Module {}

            $result = Test-ADModule

            $result | Should -Be $true
            Assert-MockCalled Get-Module -Times 1 -ParameterFilter { $ListAvailable -and $Name -eq 'ActiveDirectory' }
            Assert-MockCalled Get-Module -Times 1 -ParameterFilter { $Name -eq 'ActiveDirectory' -and -not $ListAvailable }
            Assert-MockCalled Import-Module -Times 1 -ParameterFilter { $Name -eq 'ActiveDirectory' }
        }
    }

    Context "When ActiveDirectory module is available, but Import-Module fails" {
        It "should return false and write an error" {
            Mock Get-Module { return @{ Name = 'ActiveDirectory' } } -ParameterFilter { $ListAvailable -and $Name -eq 'ActiveDirectory' }
            Mock Get-Module { return $null } -ParameterFilter { $Name -eq 'ActiveDirectory' -and -not $ListAvailable }
            Mock Import-Module { throw "Access denied" }
            Mock Write-Error {} # Suppress error output during test

            $result = Test-ADModule

            $result | Should -Be $false
            Assert-MockCalled Get-Module -Times 1 -ParameterFilter { $ListAvailable -and $Name -eq 'ActiveDirectory' }
            Assert-MockCalled Get-Module -Times 1 -ParameterFilter { $Name -eq 'ActiveDirectory' -and -not $ListAvailable }
            Assert-MockCalled Import-Module -Times 1 -ParameterFilter { $Name -eq 'ActiveDirectory' }
            Assert-MockCalled Write-Error -Times 1
        }
    }
}
