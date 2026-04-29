BeforeAll {
    # Wir laden die Funktion als Text und führen sie aus, ohne den restlichen Code von MultiRDP.ps1 auszuführen.
    # Da das Skript Set-Culture enthält (welches auf manchen non-Windows Umgebungen nicht verfügbar ist),
    # und wir nur Test-RDPConnection isoliert testen wollen, ist das hier die sicherste Lösung.
    $scriptContent = Get-Content -Path "$PSScriptRoot/../MultiRDP.ps1" -Raw

    # Parse the file via AST instead of regex for reliable extraction
    $ast = [System.Management.Automation.Language.Parser]::ParseInput($scriptContent, [ref]$null, [ref]$null)
    $functionAst = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $args[0].Name -eq 'Test-RDPConnection'}, $true)

    if ($functionAst) {
        $functionDefinition = $functionAst[0].Extent.Text
        Invoke-Expression $functionDefinition
    } else {
        throw "Could not find Test-RDPConnection in MultiRDP.ps1"
    }
}

Describe "Test-RDPConnection" {
    Context "When testing a responsive host" {
        It "Should correctly resolve DNS and Ping" {
            $rawResult = Test-RDPConnection -Computers "localhost"

            # The function outputs Format-Table data in addition to the result object.
            # Filter the return array to get just the PSCustomObject containing the actual results.
            $result = $rawResult | Where-Object { $_ -is [PSCustomObject] -and $null -ne $_.ComputerName }

            $result | Should -Not -BeNullOrEmpty
            if ($result -is [array]) {
                $result.Count | Should -Be 1
                $result[0].ComputerName | Should -Be "localhost"
                $result[0].DNSStatus | Should -Be "Aufgelöst"
                $result[0].PingStatus | Should -Be "Erreichbar"
            } else {
                $result.ComputerName | Should -Be "localhost"
                $result.DNSStatus | Should -Be "Aufgelöst"
                $result.PingStatus | Should -Be "Erreichbar"
            }
        }
    }

    Context "When testing an unresponsive host" {
        It "Should report 'Nicht aufgelöst' and 'Nicht erreichbar' for a non-existent host" {
            $invalidHost = "this.host.does.not.exist.internal.123456"
            $rawResult = Test-RDPConnection -Computers $invalidHost

            $result = $rawResult | Where-Object { $_ -is [PSCustomObject] -and $null -ne $_.ComputerName }

            $result | Should -Not -BeNullOrEmpty
            if ($result -is [array]) {
                $result.Count | Should -Be 1
                $result[0].ComputerName | Should -Be $invalidHost
                $result[0].DNSStatus | Should -Be "Nicht aufgelöst"
                $result[0].PingStatus | Should -Be "Nicht erreichbar"
                $result[0].RDPPortStatus | Should -Be "Geschlossen"
            } else {
                $result.ComputerName | Should -Be $invalidHost
                $result.DNSStatus | Should -Be "Nicht aufgelöst"
                $result.PingStatus | Should -Be "Nicht erreichbar"
                $result.RDPPortStatus | Should -Be "Geschlossen"
            }
        }
    }

    Context "When testing multiple hosts" {
        It "Should return a result for each provided host" {
            $hosts = @("localhost", "another.invalid.host.xyz")
            $rawResult = Test-RDPConnection -Computers $hosts

            $results = $rawResult | Where-Object { $_ -is [PSCustomObject] -and $null -ne $_.ComputerName }

            $results | Should -Not -BeNullOrEmpty

            $results.Count | Should -Be 2

            $localhostResult = $results | Where-Object { $_.ComputerName -eq "localhost" }
            $invalidResult = $results | Where-Object { $_.ComputerName -eq "another.invalid.host.xyz" }

            $localhostResult | Should -Not -BeNullOrEmpty
            $invalidResult | Should -Not -BeNullOrEmpty

            $localhostResult.DNSStatus | Should -Be "Aufgelöst"
            $invalidResult.DNSStatus | Should -Be "Nicht aufgelöst"
        }
    }
}
