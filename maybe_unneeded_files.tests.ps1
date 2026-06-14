# Pester tests for maybe_unneeded_files.ps1

$scriptPath = Join-Path $PSScriptRoot "maybe_unneeded_files.ps1"
. $scriptPath

Describe "New-HTMLReport" {
    BeforeAll {
        $TestDrive = "TestDrive:\"
        $script:OutputDirectory = $TestDrive
    }

    It "Generates an HTML file with correct title" {
        $testFilePath = Join-Path $TestDrive "test_input.txt"
        Set-Content -Path $testFilePath -Value "C:\test\file1.txt"

        $reportPath = New-HTMLReport -Title "TestReportTitle" -FilePaths @($testFilePath)

        $reportPath | Should -Exist
        $content = Get-Content $reportPath -Raw
        $content | Should -Match "<title>File Analysis - TestReportTitle</title>"
        $content | Should -Match "<h1>TestReportTitle</h1>"
    }

    It "Encodes file paths properly and adds archive class if needed" {
        $testFilePath = Join-Path $TestDrive "test_input2.txt"
        Set-Content -Path $testFilePath -Value "C:\test\file2.txt", "C:\test\---ARCHIVE---\archived.txt", "C:\test\special&file.txt"

        $reportPath = New-HTMLReport -Title "EncodingTest" -FilePaths @($testFilePath)

        $content = Get-Content $reportPath -Raw
        $content | Should -Match "C:\\test\\file2.txt"
        # Should contain the archive class for the second file
        $content | Should -Match "class=`"file-item archive`""
        $content | Should -Match "C:\\test\\---ARCHIVE---\\archived.txt"
        # The & should be HTML encoded as &amp;
        $content | Should -Match "C:\\test\\special&amp;file.txt"
    }

    It "Handles missing input files gracefully" {
        $missingFile = Join-Path $TestDrive "does_not_exist.txt"

        $reportPath = New-HTMLReport -Title "MissingFileTest" -FilePaths @($missingFile)

        $reportPath | Should -Exist
        $content = Get-Content $reportPath -Raw
        # Should still generate the basic structure
        $content | Should -Match "MissingFileTest"
    }
}
