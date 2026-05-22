$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'

# Mock Import-Module to prevent it from failing when ActiveDirectory isn't available
BeforeAll {
    Mock Import-Module {}
    . "$here\$sut"
}

Describe 'Compare-GroupMembership' {
    It 'Returns false when TemplateGroups is null or empty' {
        $result = Compare-GroupMembership -TemplateGroups $null -UserGroups @("CN=Group1,OU=Test,DC=domain,DC=com")
        $result | Should -Be $false
    }

    It 'Returns false when UserGroups is null or empty' {
        $result = Compare-GroupMembership -TemplateGroups @("CN=Group1,OU=Test,DC=domain,DC=com") -UserGroups $null
        $result | Should -Be $false
    }

    It 'Returns true for exact matches' {
        $templateGroups = @("CN=AdminGroup,OU=Groups,DC=domain,DC=com")
        $userGroups = @("CN=AdminGroup,OU=Groups,DC=domain,DC=com")

        $result = Compare-GroupMembership -TemplateGroups $templateGroups -UserGroups $userGroups
        $result | Should -Be $true
    }

    It 'Returns true when CN matches but OU differs' {
        $templateGroups = @("CN=SalesGroup,OU=Global,DC=domain,DC=com")
        $userGroups = @("CN=SalesGroup,OU=Local,DC=domain,DC=com")

        $result = Compare-GroupMembership -TemplateGroups $templateGroups -UserGroups $userGroups
        $result | Should -Be $true
    }

    It 'Returns false for completely unmatched groups' {
        $templateGroups = @("CN=HRGroup,OU=Groups,DC=domain,DC=com")
        $userGroups = @("CN=ITGroup,OU=Groups,DC=domain,DC=com")

        $result = Compare-GroupMembership -TemplateGroups $templateGroups -UserGroups $userGroups
        $result | Should -Be $false
    }

    It 'Returns true when there is at least one common group among many' {
        $templateGroups = @(
            "CN=HRGroup,OU=Groups,DC=domain,DC=com",
            "CN=CommonGroup,OU=Groups,DC=domain,DC=com"
        )
        $userGroups = @(
            "CN=ITGroup,OU=Groups,DC=domain,DC=com",
            "CN=CommonGroup,OU=OtherGroups,DC=domain,DC=com"
        )

        $result = Compare-GroupMembership -TemplateGroups $templateGroups -UserGroups $userGroups
        $result | Should -Be $true
    }

    It 'Handles group strings that do not contain a comma properly' {
        $templateGroups = @("SimpleGroup1", "SimpleGroup2")
        $userGroups = @("SimpleGroup3", "SimpleGroup1")

        $result = Compare-GroupMembership -TemplateGroups $templateGroups -UserGroups $userGroups
        $result | Should -Be $true
    }
}
