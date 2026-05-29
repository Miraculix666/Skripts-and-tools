BeforeAll {
    Mock Import-Module {} -ParameterFilter { $Name -eq 'ActiveDirectory' }
    $scriptContent = Get-Content -Path "$PSScriptRoot/CopyCreateNewUser.ps1" -Raw
    $functionRegex = '(?s)function Create-LogDirectory \{.*?\n\}'
    if ($scriptContent -match $functionRegex) {
        Invoke-Expression $matches[0]
    }
}

Describe "Create-LogDirectory" {
    Context "When directory does not exist" {
        It "Creates the directory and returns the path using TestDrive" {
            $testDrivePath = "TestDrive:\ADUserCreationLogs"
            if (Test-Path $testDrivePath) {
                Remove-Item -Path $testDrivePath -Recurse -Force
            }
            $result = Create-LogDirectory -logDir $testDrivePath
            $result | Should -Be $testDrivePath
            (Test-Path $testDrivePath) | Should -Be $true
        }
    }

    Context "When directory already exists" {
        It "Does not attempt to recreate the directory and returns the path" {
            $testDrivePath = "TestDrive:\ExistingLogs"
            New-Item -ItemType Directory -Path $testDrivePath -Force | Out-Null
            Mock New-Item { }
            $result = Create-LogDirectory -logDir $testDrivePath
            $result | Should -Be $testDrivePath
            Assert-MockCalled New-Item -Times 0
        }
    }
}
