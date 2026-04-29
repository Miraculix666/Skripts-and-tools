Describe "Add-UserReportEntry" {
    BeforeAll {
        # Using PowerShell's AST to safely extract the function
        $ast = [System.Management.Automation.Language.Parser]::ParseInput((Get-Content "Enhanced-AD-User-Managment.ps1" -Raw), [ref]$null, [ref]$null)
        $functionAst = $ast.Find({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Add-UserReportEntry'
        }, $true)

        if ($functionAst) {
            # Execute the function definition to load it into scope
            Invoke-Expression $functionAst.Extent.Text
        } else {
            throw "Function Add-UserReportEntry not found"
        }
    }

    Context "When global report data is initialized" {
        BeforeEach {
            $global:userReportData = [System.Collections.Generic.List[PSObject]]::new()
        }

        AfterEach {
            $global:userReportData = $null
        }

        It "Should add a new entry to the report" {
            Add-UserReportEntry -SamAccountName "testuser" -Status "Erstellt" -Detail "Success"

            $global:userReportData.Count | Should -Be 1
            $global:userReportData[0].SamAccountName | Should -Be "testuser"
            $global:userReportData[0].Status | Should -Be "Erstellt"
            $global:userReportData[0].Detail | Should -Be "Success"
            $global:userReportData[0].Timestamp | Should -Not -BeNullOrEmpty
        }
    }

    Context "When global report data is not initialized" {
        BeforeEach {
            $global:userReportData = $null
        }

        It "Should not throw an error and ignore the entry" {
            $action = { Add-UserReportEntry -SamAccountName "testuser" -Status "Erstellt" -Detail "Success" }

            $action | Should -Not -Throw
            $global:userReportData | Should -BeNullOrEmpty
        }
    }

    Context "When detail is not provided" {
        BeforeEach {
            $global:userReportData = [System.Collections.Generic.List[PSObject]]::new()
        }

        AfterEach {
            $global:userReportData = $null
        }

        It "Should use an empty string for detail" {
            Add-UserReportEntry -SamAccountName "testuser" -Status "Erstellt"

            $global:userReportData.Count | Should -Be 1
            $global:userReportData[0].SamAccountName | Should -Be "testuser"
            $global:userReportData[0].Status | Should -Be "Erstellt"
            $global:userReportData[0].Detail | Should -Be ""
        }
    }
}
