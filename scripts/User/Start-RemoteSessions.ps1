#
# FILE: Start-RemoteSessions.ps1
# DESCRIPTION: Opens multiple remote sessions and tiles them on the second monitor.
#
# AI-DISCOVERED SOURCES:
# - ChatGPT
# - Microsoft Docs (Get-CimInstance, Write-Verbose, Read-Host, Hashtables, System.Windows.Forms.Screen)
# - Stack Overflow (PowerShell RDP window positioning, cmdkey usage, Test-NetConnection)
#
# USER-PROVIDED SOURCES:
# - None
#
# VERSION: 2.2.0
# DATE: 2025-09-05
# AUTHOR: PS-Coding
#
<#
.SYNOPSIS
    Starts multiple RDP sessions, tiles them reliably on the second monitor, and handles credentials securely.

.DESCRIPTION
    This script reads a list of computer names, performs a connectivity check (ping and RDP port),
    and then opens a Remote Desktop Connection for each reachable computer.

    By default, it uses the currently logged-on user's credentials for a seamless single sign-on experience.
    If specific credentials are needed, they can be provided using the -UserName and -Password parameters.
    The script securely handles these credentials using the Windows Credential Manager (cmdkey.exe).

    A summary of successful and failed connections is provided at the end. Verbose output is enabled by default.

.PARAMETER ComputerName
    A comma-separated list of computer names to connect to.

.PARAMETER InputFile
    The full path to a text (.txt) or CSV (.csv) file with computer names.

.PARAMETER Preset
    The name of a predefined internal computer list to use.

.PARAMETER UserName
    The username for the RDP connections. If not provided, the current user is used.

.PARAMETER Password
    The password (SecureString) for the specified user. Prompts if -UserName is used without a password.

.EXAMPLE
    .\Start-RemoteSessions.ps1 -Preset "R204"

.EXAMPLE
    .\Start-RemoteSessions.ps1 -InputFile C:\temp\computers.txt -UserName "domain\adminuser"
#>
[CmdletBinding()]
param (
    [Parameter(Position = 0)]
    [string[]]$ComputerName = @(),

    [Parameter(Position = 1)]
    [string]$InputFile,

    [Parameter(Position = 2)]
    [string]$Preset,

    [Parameter(Position = 3)]
    [string]$UserName,

    [Parameter(Position = 4)]
    [System.Security.SecureString]$Password
)

# --- SCRIPT-START ---

# Set German locale and force Verbose output
Set-Culture -Culture de-DE
$VerbosePreference = "Continue"

Write-Verbose "Starting script execution on $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')."

#region Variable Definitions
$InternalLists = @{
    "Default"      = @(
        "Server01",
        "Server02",
        "Server03"
    );
    "R202"         = @(
        "C26PBZE70004200", "C26PBZE70004201", "C26PBZE70004202", "C26PBZE70004203",
        "C26PBZE70004204", "C26PBZE70004205", "C26PBZE70004206", "C26PBZE70004207",
        "C26PBZE70004208", "C26PBZE70004209", "C26PBZE70004210", "C26PBZE70004211",
        "C26PBZE70004212", "C26PBZE70004299"
    );
    "R204"         = @( # Formerly R209
        "C26PBZE70004400", "C26PBZE70004401", "C26PBZE70004402", "C26PBZE70004403",
        "C26PBZE70004404", "C26PBZE70004405", "C26PBZE70004406", "C26PBZE70004407",
        "C26PBZE70004408", "C26PBZE70004409", "C26PBZE70004410", "C26PBZE70004411",
        "C26PBZE70004412", "C26PBZE70004499"
    );
    "Server"       = @(
        "web-server01",
        "sql-server02"
    );
    "Testumgebung" = @(
        "dev-vm-01",
        "dev-vm-02",
        "test-vm-01"
    )
}
#endregion

#region Get Computer List
$ComputersToConnect = @()

if ($InputFile) {
    Write-Verbose "Reading computer names from '$InputFile'."
    if (-not (Test-Path -Path $InputFile)) { Write-Error "File not found: '$InputFile'"; return }
    $ext = [System.IO.Path]::GetExtension($InputFile).ToLower()
    if ($ext -eq ".csv") {
        try { $ComputersToConnect = Import-Csv -Path $InputFile -Delimiter ';' | Select-Object -ExpandProperty ComputerName }
        catch { Write-Error "Failed to read CSV. Ensure it has a 'ComputerName' column."; return }
    }
    elseif ($ext -eq ".txt") { $ComputersToConnect = Get-Content -Path $InputFile | Where-Object { $_.Trim() -ne "" } }
    else { Write-Error "Unsupported file format. Use .txt or .csv."; return }
}
elseif ($ComputerName.Count -gt 0) { $ComputersToConnect = $ComputerName }
elseif ($Preset) {
    if ($InternalLists.ContainsKey($Preset)) { $ComputersToConnect = $InternalLists[$Preset] }
    else { Write-Error "Preset '$Preset' not found. Available: $($InternalLists.Keys -join ', ')"; return }
}
else { $ComputersToConnect = $InternalLists["Default"] }

if ($ComputersToConnect.Count -eq 0) { Write-Warning "No computers to connect to. Exiting."; return }
#endregion

#region Handle Credentials
$useCurrentUserCredentials = $false
if (-not $UserName) {
    $useCurrentUserCredentials = $true
    $UserName = "$($env:USERDOMAIN)\$($env:USERNAME)"
    Write-Verbose "Using current user for SSO: '$UserName'."
}
elseif (-not $Password) {
    Write-Host "Please enter the password for '$UserName':" -NoNewline
    $Password = Read-Host -AsSecureString
    if ($Password.Length -eq 0) { Write-Error "Password cannot be empty. Exiting."; return }
}
#endregion

#region Get Screen Dimensions for Tiling
$useTiling = $false
try {
    Add-Type -AssemblyName System.Windows.Forms
    $screens = [System.Windows.Forms.Screen]::AllScreens
    if ($screens.Count -gt 1) {
        $secondaryScreen = $screens | Where-Object { -not $_.Primary } | Select-Object -First 1
        if ($secondaryScreen) {
            $useTiling = $true
            $bounds = $secondaryScreen.WorkingArea
            $xOffset = $bounds.Left
            $yOffset = $bounds.Top
            $screenWidth = $bounds.Width
            $screenHeight = $bounds.Height
            Write-Verbose "Secondary screen detected for tiling at ${xOffset},${yOffset} with resolution ${screenWidth}x${screenHeight}."
        }
    }
}
catch { Write-Warning "Could not get screen info. Tiling disabled. Error: $($_.Exception.Message)" }
#endregion

#region Create and Start RDP Sessions
if ($useTiling) {
    $numComputers = $ComputersToConnect.Count
    $numRows = [System.Math]::Ceiling([System.Math]::Sqrt($numComputers))
    $numCols = [System.Math]::Ceiling($numComputers / $numRows)
    $tileWidth = [int]($screenWidth / $numCols)
    $tileHeight = [int]($screenHeight / $numRows)
    Write-Verbose "Tiling grid: $numCols columns x $numRows rows. Tile size: ${tileWidth}x${tileHeight}."
}

$successList = @()
$failureList = @{}
$counter = 0

foreach ($computer in $ComputersToConnect) {
    $counter++
    Write-Verbose "($counter/$($ComputersToConnect.Count)) Processing '$computer'..."
    
    # 1. Pre-connection check
    Write-Verbose "  Checking connectivity to '$computer'..."
    $connectionTest = Test-NetConnection -ComputerName $computer -Port 3389 -WarningAction SilentlyContinue
    if (-not $connectionTest.TcpTestSucceeded) {
        Write-Warning "  Host '$computer' is not reachable on port 3389. Skipping."
        $failureList[$computer] = "Not Reachable"
        continue
    }
    Write-Verbose "  Connectivity check successful."

    $plainTextPassword = ""
    $tempRDPFile = Join-Path -Path $env:TEMP -ChildPath "$($computer)-RDP-$(Get-Random).rdp"

    try {
        # 2. Handle credentials if provided
        if (-not $useCurrentUserCredentials) {
            $plainTextPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
                [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
            )
            Write-Verbose "  Temporarily adding credentials for '$computer'..."
            cmdkey.exe /add:"TERMSRV/$computer" /user:"$UserName" /pass:"$plainTextPassword"
            if ($LASTEXITCODE -ne 0) {
                Write-Error "  Failed to store credentials for '$computer' using cmdkey.exe. Check permissions or group policies."
                $failureList[$computer] = "Credential Error"
                continue
            }
        }
        
        # 3. Create RDP file with robust settings
        $rdpContent = @"
full address:s:$computer
enablecredsspsupport:i:1
authentication level:i:2
promptcredentialonce:i:0
negotiate security layer:i:1
redirectclipboard:i:1
"@
        
        if ($useTiling) {
            $row = [int](($counter - 1) / $numCols)
            $col = [int](($counter - 1) % $numCols)
            $xPos = $xOffset + ($col * $tileWidth)
            $yPos = $yOffset + ($row * $tileHeight)
            
            $rdpContent += "`ndesktopwidth:i:$tileWidth`ndesktopheight:i:$tileHeight"
            $rdpContent += "`nwinposstr:s:0,1,$xPos,$yPos,$($xPos + $tileWidth),$($yPos + $tileHeight)"
        } else {
            $rdpContent += "`ndesktopwidth:i:1280`ndesktopheight:i:800"
        }

        Set-Content -Path $tempRDPFile -Value $rdpContent -Encoding UTF8
        
        # 4. Launch RDP session
        Write-Verbose "  Launching RDP session for '$computer'..."
        Start-Process -FilePath "mstsc.exe" -ArgumentList "`"$tempRDPFile`""
        $successList += $computer
        Start-Sleep -Milliseconds 500
    }
    catch {
        Write-Warning "  An error occurred while launching session for '$computer': $($_.Exception.Message)"
        $failureList[$computer] = "Launch Error"
    }
    finally {
        # 5. Clean up
        if (-not $useCurrentUserCredentials -and $plainTextPassword) {
            Write-Verbose "  Removing temporary credentials for '$computer'."
            cmdkey.exe /delete:"TERMSRV/$computer"
            $plainTextPassword = $null
        }
        if (Test-Path $tempRDPFile) { Remove-Item $tempRDPFile -Force -ErrorAction SilentlyContinue }
    }
}

# Final Summary
Write-Host "`n---"
Write-Host "Execution Summary:"
Write-Host "  User: $UserName"
Write-Host "---"
if ($successList.Count -gt 0) {
    Write-Host "Successfully launched sessions for ($($successList.Count)): " -ForegroundColor Green
    $successList | ForEach-Object { Write-Host "  - $_" }
}
if ($failureList.Count -gt 0) {
    Write-Host "Failed to launch sessions for ($($failureList.Count)): " -ForegroundColor Red
    $failureList.GetEnumerator() | ForEach-Object { Write-Host "  - $($_.Name): $($_.Value)" }
}
Write-Host "---"
Write-Verbose "Script execution finished."

