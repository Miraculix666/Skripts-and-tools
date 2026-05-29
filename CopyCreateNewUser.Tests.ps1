Describe "Export-TemplateUserData Edge Cases" {
    BeforeAll {
        $scriptPath = "$PSScriptRoot\CopyCreateNewUser.ps1"
        $scriptContent = Get-Content -Path $scriptPath -Raw

        # Extract the Export-TemplateUserData function
        $pattern = '(?s)function Export-TemplateUserData \{.*?(?=\n# Funktion zum Erstellen eines neuen AD-Benutzers)'
        if ($scriptContent -match $pattern) {
            $functionCode = $matches[0]
            Invoke-Expression $functionCode
        } else {
            throw "Could not extract function Export-TemplateUserData"
        }

        function Write-Log { param($Message) }
        function Write-LogMessage { param($Message, $Type) }
    }

    Context "When Get-ADUser returns partial user objects" {
        It "Should handle user with missing properties without error" {
            $templateUser = "PartialUser"
            $csvPath = "TestDrive:\PartialUser.csv"

            Mock Get-ADUser {
                return [PSCustomObject]@{
                    SamAccountName = "PartialUser"
                    GivenName = "John"
                    Surname = "Doe"
                    # Missing other properties like Office, Department, etc.
                }
            } -ParameterFilter { $Identity -eq $templateUser }

            Export-TemplateUserData -TemplateUser $templateUser -CsvPath $csvPath

            # Verify the exported file
            $resolvedCsvPath = (Resolve-Path $csvPath).Path
            $csvContent = Import-Csv -Path $resolvedCsvPath -Delimiter ';'

            $csvContent.SamAccountName | Should -Be "PartialUser"
            $csvContent.GivenName | Should -Be "John"
            $csvContent.Surname | Should -Be "Doe"
            $csvContent.Department | Should -BeNullOrEmpty
            $csvContent.Office | Should -BeNullOrEmpty
        }

        It "Should handle user with completely empty properties" {
            $templateUser = "EmptyUser"
            $csvPath = "TestDrive:\EmptyUser.csv"

            Mock Get-ADUser {
                return [PSCustomObject]@{
                    SamAccountName = "EmptyUser"
                }
            } -ParameterFilter { $Identity -eq $templateUser }

            Export-TemplateUserData -TemplateUser $templateUser -CsvPath $csvPath

            $resolvedCsvPath = (Resolve-Path $csvPath).Path
            $csvContent = Import-Csv -Path $resolvedCsvPath -Delimiter ';'

            $csvContent.SamAccountName | Should -Be "EmptyUser"
            $csvContent.GivenName | Should -BeNullOrEmpty
            $csvContent.Department | Should -BeNullOrEmpty
        }

        It "Should handle Get-ADUser returning `$null gracefully" {
            $templateUser = "NullUser"
            $csvPath = "TestDrive:\NullUser.csv"

            Mock Get-ADUser {
                return $null
            } -ParameterFilter { $Identity -eq $templateUser }

            { Export-TemplateUserData -TemplateUser $templateUser -CsvPath $csvPath } | Should -Not -Throw
        }
    }
}
