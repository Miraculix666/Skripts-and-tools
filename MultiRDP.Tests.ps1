$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
$scriptPath = Join-Path $here $sut

# Dot-source the script but prevent it from executing its main logic.
# The script checks $InputFile, $ComputerName, and $Preset. We provide a non-existent Preset.
# It warns and exits because no computers to connect to are found.
. $scriptPath -Preset "TestOnlyNonExistentPreset" -ErrorAction SilentlyContinue | Out-Null

Describe "Test-RDPConnection" {
    Context "Connectivity Tests" {
        BeforeEach {
            # Mock job management cmdlets to run synchronously and capture output
            Mock Start-Job {
                param($ScriptBlock, $ArgumentList)
                if ($ArgumentList) {
                    $result = & $ScriptBlock @ArgumentList
                } else {
                    $result = & $ScriptBlock
                }
                $mockJob = [PSCustomObject]@{
                    Id = 1
                    State = 'Completed'
                    Result = $result
                }
                return $mockJob
            }

            Mock Receive-Job {
                param($Job)
                return $Job.Result
            }

            Mock Remove-Job { return }

            # Mock UI cmdlets
            Mock Write-Progress { return }
            Mock Write-Host { return }
            Mock Format-Table { return }
        }

        It "Should report 'Erreichbar', 'Geöffnet', and 'Aufgelöst' for a fully working host" {
            # Mock ping success
            Mock Test-Connection { return $true }

            # Mock TCP client success
            $mockAsyncWaitHandle = [PSCustomObject]@{}
            $mockAsyncWaitHandle | Add-Member -MemberType ScriptMethod -Name WaitOne -Value { param($timeout, $exitContext) return $true }

            $mockTcpClient = [PSCustomObject]@{ Connected = $true }
            $mockTcpClient | Add-Member -MemberType ScriptMethod -Name BeginConnect -Value { param($host, $port, $requestCallback, $state) return [PSCustomObject]@{ AsyncWaitHandle = $mockAsyncWaitHandle } }
            $mockTcpClient | Add-Member -MemberType ScriptMethod -Name Close -Value { }
            Mock New-Object {
                return $mockTcpClient
            } -ParameterFilter { $TypeName -eq 'System.Net.Sockets.TcpClient' }

            # 'localhost' is generally resolvable by DNS
            $result = Test-RDPConnection -Computers @("localhost")

            $result.Count | Should -Be 1
            $result[0].ComputerName | Should -Be "localhost"
            $result[0].DNSStatus | Should -Be "Aufgelöst"
            $result[0].PingStatus | Should -Be "Erreichbar"
            $result[0].RDPPortStatus | Should -Be "Geöffnet"
        }

        It "Should report 'Nicht erreichbar' when Ping fails" {
            Mock Test-Connection { return $false }

            $mockAsyncWaitHandle = [PSCustomObject]@{}
            $mockAsyncWaitHandle | Add-Member -MemberType ScriptMethod -Name WaitOne -Value { param($timeout, $exitContext) return $true }

            $mockTcpClient = [PSCustomObject]@{ Connected = $true }
            $mockTcpClient | Add-Member -MemberType ScriptMethod -Name BeginConnect -Value { param($host, $port, $requestCallback, $state) return [PSCustomObject]@{ AsyncWaitHandle = $mockAsyncWaitHandle } }
            $mockTcpClient | Add-Member -MemberType ScriptMethod -Name Close -Value { }
            Mock New-Object {
                return $mockTcpClient
            } -ParameterFilter { $TypeName -eq 'System.Net.Sockets.TcpClient' }

            $result = Test-RDPConnection -Computers @("localhost")

            $result[0].PingStatus | Should -Be "Nicht erreichbar"
            $result[0].RDPPortStatus | Should -Be "Geöffnet"
        }

        It "Should report 'Geschlossen' when RDP port check times out" {
            Mock Test-Connection { return $true }

            # Mock TCP client failure (WaitOne returns false = timeout)
            $mockAsyncWaitHandle = [PSCustomObject]@{}
            $mockAsyncWaitHandle | Add-Member -MemberType ScriptMethod -Name WaitOne -Value { param($timeout, $exitContext) return $false }

            $mockTcpClient = [PSCustomObject]@{ Connected = $false }
            $mockTcpClient | Add-Member -MemberType ScriptMethod -Name BeginConnect -Value { param($host, $port, $requestCallback, $state) return [PSCustomObject]@{ AsyncWaitHandle = $mockAsyncWaitHandle } }
            $mockTcpClient | Add-Member -MemberType ScriptMethod -Name Close -Value { }
            Mock New-Object {
                return $mockTcpClient
            } -ParameterFilter { $TypeName -eq 'System.Net.Sockets.TcpClient' }

            $result = Test-RDPConnection -Computers @("localhost")

            $result[0].PingStatus | Should -Be "Erreichbar"
            $result[0].RDPPortStatus | Should -Be "Geschlossen"
        }

        It "Should report 'Nicht aufgelöst' for an unresolvable host" {
            Mock Test-Connection { return $false }
            Mock New-Object {
                throw "Mock exception"
            } -ParameterFilter { $TypeName -eq 'System.Net.Sockets.TcpClient' }

            # 'invalid.nonexistent.test.local' should fail DNS resolution
            $result = Test-RDPConnection -Computers @("invalid.nonexistent.test.local")

            $result[0].DNSStatus | Should -Be "Nicht aufgelöst"
            $result[0].PingStatus | Should -Be "Nicht erreichbar"
            $result[0].RDPPortStatus | Should -Be "Geschlossen"
        }
    }
}
