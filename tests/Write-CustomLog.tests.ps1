BeforeAll {
    # Extract just the Write-CustomLog function from the script to avoid syntax errors from the truncated script
    $content = Get-Content -Path "$PSScriptRoot/../UserCopyCreate.ps1"
    $funcText = $content[52..81] -join "`n"
    Invoke-Expression $funcText
}

Describe "Write-CustomLog" {
    BeforeEach {
        # Clean up or setup log path
        $global:LogPath = "$PSScriptRoot/test_log.log"
        if (Test-Path $global:LogPath) {
            Remove-Item $global:LogPath -Force
        }
    }

    AfterEach {
        if (Test-Path $global:LogPath) {
            Remove-Item $global:LogPath -Force
        }
    }

    It "Writes an INFO message to the log file" {
        $msg = "Test Info Message"
        Write-CustomLog -Message $msg

        $logContent = Get-Content $global:LogPath
        $logContent | Should -Match "^.*\d{2}\.\d{2}\.\d{4} \d{2}:\d{2}:\d{2}.*\[INFO\] $msg"
    }

    It "Writes a WARNUNG message to the log file" {
        $msg = "Test Warning Message"
        Write-CustomLog -Message $msg -Level "WARNUNG"

        $logContent = Get-Content $global:LogPath
        $logContent | Should -Match "^.*\d{2}\.\d{2}\.\d{4} \d{2}:\d{2}:\d{2}.*\[WARNUNG\] $msg"
    }

    It "Writes a FEHLER message to the log file" {
        $msg = "Test Error Message"
        Write-CustomLog -Message $msg -Level "FEHLER"

        $logContent = Get-Content $global:LogPath
        $logContent | Should -Match "^.*\d{2}\.\d{2}\.\d{4} \d{2}:\d{2}:\d{2}.*\[FEHLER\] $msg"
    }

    It "Writes a DEBUG message to the log file" {
        $msg = "Test Debug Message"
        Write-CustomLog -Message $msg -Level "DEBUG"

        $logContent = Get-Content $global:LogPath
        $logContent | Should -Match "^.*\d{2}\.\d{2}\.\d{4} \d{2}:\d{2}:\d{2}.*\[DEBUG\] $msg"
    }

    It "Calls Write-Host with the correct colors" {
        $msg = "Test Console Message"

        Mock Write-Host { }

        Write-CustomLog -Message $msg -Level "WARNUNG"
        Assert-MockCalled Write-Host -Times 1 -ParameterFilter { $ForegroundColor -eq "Yellow" }

        Write-CustomLog -Message $msg -Level "FEHLER"
        Assert-MockCalled Write-Host -Times 1 -ParameterFilter { $ForegroundColor -eq "Red" }
    }
}
