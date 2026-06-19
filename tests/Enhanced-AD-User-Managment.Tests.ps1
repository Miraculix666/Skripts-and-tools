$scriptPath = Join-Path $PSScriptRoot "..\Enhanced-AD-User-Managment.ps1"

# In a monolithic script with begin/process/end blocks, executing it might cause side effects.
# To safely test the function, we use the Abstract Syntax Tree (AST) to extract the function definition.
$scriptContent = Get-Content -Path $scriptPath -Raw

# Native, safe AST parsing without execution
$ast = [System.Management.Automation.Language.Parser]::ParseInput($scriptContent, [ref]$null, [ref]$null)
$functionAst = $ast.Find({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $args[0].Name -eq 'Add-UserReportEntry' }, $true)

if ($null -ne $functionAst) {
    # Extract the full function and execute it to load into scope
    Invoke-Expression $functionAst.Extent.Text
} else {
    Throw "Could not find Add-UserReportEntry in $scriptPath using AST parsing."
}

Describe "Add-UserReportEntry" {
    BeforeAll {
        # Ensure userReportData is fresh for the suite
        $global:userReportData = $null
    }

    Context "When global report array is initialized" {
        BeforeEach {
            $global:userReportData = New-Object System.Collections.ArrayList
        }

        It "should add an entry to the userReportData array" {
            Add-UserReportEntry -SamAccountName "testuser" -Status "Kopiert" -Detail "Test detail"

            $global:userReportData.Count | Should -Be 1
        }

        It "should add correct values for SamAccountName, Status, and Detail" {
            Add-UserReportEntry -SamAccountName "jdoe" -Status "Erstellt" -Detail "Created properly"

            $entry = $global:userReportData[0]
            $entry.SamAccountName | Should -Be "jdoe"
            $entry.Status | Should -Be "Erstellt"
            $entry.Detail | Should -Be "Created properly"
            $entry.Timestamp | Should -Match "^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$"
        }

        It "should handle missing Detail parameter and default to empty string" {
            Add-UserReportEntry -SamAccountName "asmith" -Status "Modifiziert"

            $entry = $global:userReportData[0]
            $entry.Detail | Should -Be ""
        }
    }

    Context "When global report array is null" {
        BeforeEach {
            $global:userReportData = $null
        }

        It "should not throw an exception" {
            { Add-UserReportEntry -SamAccountName "jdoe" -Status "Fehler" } | Should -Not -Throw
        }

        It "should not initialize the array" {
            Add-UserReportEntry -SamAccountName "jdoe" -Status "Fehler"
            $global:userReportData | Should -BeNullOrEmpty
        }
    }

    Context "Parameter Validation" {
        BeforeEach {
            $global:userReportData = New-Object System.Collections.ArrayList
        }

        It "should require SamAccountName parameter" {
            { Add-UserReportEntry -Status "Erstellt" } | Should -Throw
        }

        It "should require Status parameter" {
            { Add-UserReportEntry -SamAccountName "testuser" } | Should -Throw
        }
    }
}
