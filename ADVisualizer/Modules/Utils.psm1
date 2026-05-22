# Utility functions
function Close-ExcelProcesses {
    [CmdletBinding()]
    param()

    Get-Process -Name "excel" -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            $_.CloseMainWindow() | Out-Null
            Start-Sleep -Milliseconds 500
            if (!$_.HasExited) { $_.Kill() }
        }
        catch {
            Write-Warning "Could not close Excel process: $_"
        }
    }
    Start-Sleep -Seconds 2
}

function New-OutputDirectory {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Path
    )

    $directory = Split-Path -Parent $Path
    if (-not (Test-Path $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }
}

function Remove-ExistingFile {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (Test-Path $Path) {
        Remove-Item $Path -Force -ErrorAction Stop
    }
}

Export-ModuleMember -Function Close-ExcelProcesses, New-OutputDirectory, Remove-ExistingFile
