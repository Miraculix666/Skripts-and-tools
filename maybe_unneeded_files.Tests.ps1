# Pester Test for Find-Duplicates in maybe_unneeded_files.ps1

BeforeAll {
    # Dot-source the script to load its functions
    . "$PSScriptRoot\maybe_unneeded_files.ps1"
}

Describe "Find-Duplicates" {
    BeforeEach {
        # Setup mock test directories and global OutputDirectory
        $global:OutputDirectory = Join-Path $TestDrive "TestOutput"
        New-Item -ItemType Directory -Path $global:OutputDirectory -Force | Out-Null

        $testPath = Join-Path $TestDrive "TestFiles"
        New-Item -ItemType Directory -Path $testPath -Force | Out-Null
        $script:TestPath = $testPath
    }

    AfterEach {
        # Cleanup is handled by TestDrive, but we remove the global variable to be safe
        Remove-Variable OutputDirectory -Scope Global -ErrorAction SilentlyContinue
    }

    It "Should return 'No duplicate files found' when there are no duplicates" {
        # Arrange
        New-Item -ItemType File -Path (Join-Path $script:TestPath "file1.txt") -Value "content 1" | Out-Null
        New-Item -ItemType File -Path (Join-Path $script:TestPath "file2.txt") -Value "content 2" | Out-Null

        # Act
        $resultPath = Find-Duplicates -Path $script:TestPath

        # Assert
        $resultPath | Should -Not -BeNullOrEmpty
        $content = Get-Content $resultPath
        $content | Should -Contain "No duplicate files found."
    }

    It "Should find duplicates and list them in the output file" {
        # Arrange
        # Create a file
        New-Item -ItemType File -Path (Join-Path $script:TestPath "duplicate.txt") -Value "content" | Out-Null

        # Create a subdirectory with a file of the same name
        $subDir = Join-Path $script:TestPath "SubDir"
        New-Item -ItemType Directory -Path $subDir | Out-Null
        New-Item -ItemType File -Path (Join-Path $subDir "duplicate.txt") -Value "different content" | Out-Null

        # Act
        $resultPath = Find-Duplicates -Path $script:TestPath

        # Assert
        $resultPath | Should -Not -BeNullOrEmpty
        $content = Get-Content $resultPath
        $content | Should -Match "Duplicate set: duplicate.txt"

        # Check that both file paths are listed in the file
        $content -match [regex]::Escape((Join-Path $script:TestPath "duplicate.txt")) | Should -Not -BeNullOrEmpty
        $content -match [regex]::Escape((Join-Path $subDir "duplicate.txt")) | Should -Not -BeNullOrEmpty
    }

    It "Should throw an error if Get-ChildItem fails" {
        # Arrange
        Mock Get-ChildItem { throw "Mock Error" }

        # Act & Assert
        { Find-Duplicates -Path $script:TestPath } | Should -Throw "Mock Error"
    }
}
