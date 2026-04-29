BeforeAll {
    $scriptContent = Get-Content -Raw "$PSScriptRoot/maybe_unneeded_files.ps1"

    # Use PowerShell's AST parser to safely extract the function
    $ast = [System.Management.Automation.Language.Parser]::ParseInput($scriptContent, [ref]$null, [ref]$null)
    $functionAst = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $args[0].Name -eq "Format-FileSize" }, $true)[0]

    if ($functionAst) {
        Invoke-Expression $functionAst.Extent.Text
    } else {
        throw "Could not find Format-FileSize function in maybe_unneeded_files.ps1"
    }
}

Describe "Format-FileSize" {
    It "Should format Terabytes correctly" {
        $expected = "{0:N2} TB" -f (2TB / 1TB)
        Format-FileSize -Size 2TB | Should -Be $expected
    }

    It "Should format Gigabytes correctly" {
        $expected = "{0:N2} GB" -f (1.5GB / 1GB)
        Format-FileSize -Size 1.5GB | Should -Be $expected
    }

    It "Should format Megabytes correctly" {
        $expected = "{0:N2} MB" -f (500MB / 1MB)
        Format-FileSize -Size 500MB | Should -Be $expected
    }

    It "Should format Kilobytes correctly" {
        $expected = "{0:N2} KB" -f (1500 / 1KB)
        Format-FileSize -Size 1500 | Should -Be $expected
    }

    It "Should format Bytes correctly" {
        Format-FileSize -Size 500 | Should -Be "500 B"
    }

    It "Should handle exactly 1TB" {
        $expected = "{0:N2} GB" -f (1TB / 1GB)
        Format-FileSize -Size 1TB | Should -Be $expected
    }

    It "Should handle exactly 1GB" {
        $expected = "{0:N2} MB" -f (1GB / 1MB)
        Format-FileSize -Size 1GB | Should -Be $expected
    }

    It "Should handle exactly 1MB" {
        $expected = "{0:N2} KB" -f (1MB / 1KB)
        Format-FileSize -Size 1MB | Should -Be $expected
    }

    It "Should handle exactly 1KB" {
        Format-FileSize -Size 1KB | Should -Be "1024 B"
    }

    It "Should handle 0 bytes" {
        Format-FileSize -Size 0 | Should -Be "0 B"
    }
}
