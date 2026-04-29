BeforeAll {
    $sut = "$PSScriptRoot/../UserCopyCreate.ps1"

    # Extract just the function using AST since the full script has a syntax error
    $errors = $null
    $tokens = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($sut, [ref]$tokens, [ref]$errors)
    $functionAst = $ast.Find({
        param($astNode) $astNode -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $astNode.Name -eq 'Compare-GroupMembership'
    }, $true)

    Invoke-Expression $functionAst.Extent.Text
}

Describe "Compare-GroupMembership" {
    Context "When both arrays have matching groups" {
        It "Returns true for exact match" {
            $template = @("CN=Group1,OU=Groups,DC=domain,DC=local")
            $user = @("CN=Group1,OU=Groups,DC=domain,DC=local")
            $result = Compare-GroupMembership -TemplateGroups $template -UserGroups $user
            $result | Should -Be $true
        }

        It "Returns true when only the first part matches (Common Name)" {
            $template = @("CN=Group1,OU=Different1,DC=domain,DC=local")
            $user = @("CN=Group1,OU=Different2,DC=domain,DC=local")
            $result = Compare-GroupMembership -TemplateGroups $template -UserGroups $user
            $result | Should -Be $true
        }

        It "Returns true when user has all template groups plus others" {
            $template = @("CN=Group1,DC=domain")
            $user = @("CN=Group1,DC=domain", "CN=Group2,DC=domain")
            $result = Compare-GroupMembership -TemplateGroups $template -UserGroups $user
            $result | Should -Be $true
        }

        It "Returns true when user has some of template groups" {
            $template = @("CN=Group1,DC=domain", "CN=Group2,DC=domain")
            $user = @("CN=Group1,DC=domain")
            $result = Compare-GroupMembership -TemplateGroups $template -UserGroups $user
            $result | Should -Be $true
        }
    }

    Context "When arrays have no matching groups" {
        It "Returns false" {
            $template = @("CN=Group1,DC=domain")
            $user = @("CN=Group2,DC=domain")
            $result = Compare-GroupMembership -TemplateGroups $template -UserGroups $user
            $result | Should -Be $false
        }
    }

    Context "When handling null or empty inputs" {
        It "Returns false if TemplateGroups is null" {
            $user = @("CN=Group1,DC=domain")
            $result = Compare-GroupMembership -TemplateGroups $null -UserGroups $user
            $result | Should -Be $false
        }

        It "Returns false if UserGroups is null" {
            $template = @("CN=Group1,DC=domain")
            $result = Compare-GroupMembership -TemplateGroups $template -UserGroups $null
            $result | Should -Be $false
        }

        It "Returns false if both are null" {
            $result = Compare-GroupMembership -TemplateGroups $null -UserGroups $null
            $result | Should -Be $false
        }

        It "Returns false if both are empty" {
            $result = Compare-GroupMembership -TemplateGroups @() -UserGroups @()
            $result | Should -Be $false
        }
    }
}
