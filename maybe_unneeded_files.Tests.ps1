$ErrorActionPreference = 'Stop'

Describe "Find-Duplicates" {
    BeforeAll {
        $script:testDir = Join-Path $env:TEMP "TestFileAnalysis_$([guid]::NewGuid())"
        $script:outputDir = Join-Path $script:testDir "Output"
        New-Item -ItemType Directory -Path $script:testDir -Force | Out-Null
        New-Item -ItemType Directory -Path $script:outputDir -Force | Out-Null

        # Save original location
        $originalLocation = Get-Location

        $targetScript = if ($PSScriptRoot) { Join-Path $PSScriptRoot "maybe_unneeded_files.ps1" } else { Join-Path (Get-Location) "maybe_unneeded_files.ps1" }

        try {
            Set-Location $script:testDir
            # Dot-source the script inside the temp directory to prevent global execution effects
            . $targetScript -TargetPath $script:testDir -OutputDirectory $script:outputDir
        } finally {
            Set-Location $originalLocation
        }

        # Make sure the variables are accessible to functions
        $script:OutputDirectory = $script:outputDir
    }

    AfterAll {
        if (Test-Path $script:testDir) {
            Remove-Item -Path $script:testDir -Recurse -Force
        }
    }

    Context "When there are no duplicate files" {
        It "Returns the output path and writes 'No duplicate files found.'" {
            # Arrange
            $testPath = Join-Path $script:testDir "NoDuplicates"
            New-Item -ItemType Directory -Path $testPath -Force | Out-Null
            New-Item -ItemType File -Path (Join-Path $testPath "file1.txt") -Value "content1" | Out-Null
            New-Item -ItemType File -Path (Join-Path $testPath "file2.txt") -Value "content2" | Out-Null

            # Act
            $result = Find-Duplicates -Path $testPath

            # Assert
            $result | Should -Be "$script:outputDir\Duplicates.txt"
            $content = Get-Content $result
            $content | Should -Contain "No duplicate files found."
        }
    }

    Context "When there are duplicate files" {
        It "Identifies duplicate names and logs their details" {
            # Arrange
            $testPath = Join-Path $script:testDir "WithDuplicates"
            $subDir1 = Join-Path $testPath "Dir1"
            $subDir2 = Join-Path $testPath "Dir2"
            New-Item -ItemType Directory -Path $subDir1 -Force | Out-Null
            New-Item -ItemType Directory -Path $subDir2 -Force | Out-Null

            # Create duplicate files (same name, different folders)
            New-Item -ItemType File -Path (Join-Path $subDir1 "duplicate.txt") -Value "content A" | Out-Null
            New-Item -ItemType File -Path (Join-Path $subDir2 "duplicate.txt") -Value "content B" | Out-Null
            New-Item -ItemType File -Path (Join-Path $testPath "unique.txt") -Value "content C" | Out-Null

            # Act
            $result = Find-Duplicates -Path $testPath

            # Assert
            $result | Should -Be "$script:outputDir\Duplicates.txt"
            $content = Get-Content $result -Raw
            $content | Should -Match "Duplicate set: duplicate.txt"
            $content | Should -Not -Match "unique.txt"
        }
    }

    Context "When an error occurs" {
        It "Catches the exception and throws" {
            # Arrange
            $invalidPath = "C:\Invalid\Path\That\Does\Not\Exist\Random12345"

            # Act & Assert
            { Find-Duplicates -Path $invalidPath } | Should -Throw
        }
    }
}
