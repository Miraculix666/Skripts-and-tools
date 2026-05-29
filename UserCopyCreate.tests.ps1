BeforeAll {
    # Extract the function content from line 85 to 98
    $funcLines = (Get-Content -Path "$PSScriptRoot/UserCopyCreate.ps1")[85..98]
    ${Function:Get-ValidADUser} = [ScriptBlock]::Create(($funcLines -join "
"))

    # Mock Write-CustomLog
    function Write-CustomLog { param($Message, $Level) }
    function Write-Verbose { param($Message) }
    function Get-ADUser { param($Identity, $Properties, $ErrorAction) }
}

Describe "Get-ValidADUser" {
    It "Should call Get-ADUser successfully and return the user object" {
        Mock -CommandName Get-ADUser -MockWith { return @{ Name = "testuser" } }

        $result = Get-ValidADUser -Identity "testuser"

        $result.Name | Should -Be "testuser"
        Assert-MockCalled -CommandName Get-ADUser -Times 1 -Exactly
    }

    It "Should mock Get-ADUser and throw an error, returning null and logging error" {
        Mock -CommandName Get-ADUser -MockWith { throw "AD Error" }
        Mock -CommandName Write-CustomLog

        $result = Get-ValidADUser -Identity "testuser" -Operation "TestOp"

        $result | Should -BeNullOrEmpty
        Assert-MockCalled -CommandName Get-ADUser -Times 1 -Exactly
        Assert-MockCalled -CommandName Write-CustomLog -Times 1 -Exactly -ParameterFilter { $Message -like "*Fehler bei TestOp für Benutzer 'testuser'*" -and $Level -eq "FEHLER" }
    }
}
