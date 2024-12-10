[CmdletBinding()]
param (
    [Parameter()]
    [string]$OutputPath = "C:\daten\AD_Benutzer_Gruppen_L.csv",
    [Parameter()]
    [string]$ExcelPath = "C:\daten\AD_Benutzer_Gruppen_L.xlsx"
)

function Close-ExcelProcesses {
    Get-Process -Name "excel" -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            $_.CloseMainWindow() | Out-Null
            Start-Sleep -Milliseconds 500
            if (!$_.HasExited) {
                $_.Kill()
            }
        } catch {
            Write-Warning "Could not close Excel process: $_"
        }
    }
    Start-Sleep -Seconds 2
}

function Get-ADUserGroupData {
    param (
        [string]$SearchPattern = "L*"
    )
    
    $users = Get-ADUser -Filter "SamAccountName -like '$SearchPattern'" -Properties SamAccountName, Name, MemberOf, DistinguishedName, Comment
    
    if (-not $users) {
        Write-Warning "No users found with SamAccountName starting with 'L'"
        return $null
    }
    
    $userData = @()
    $groupColors = @{}
    $colorIndex = 35  # Starting with lighter colors
    
    foreach ($user in $users) {
        Write-Verbose "Processing user: $($user.SamAccountName)"
        
        # Extract OU and numeric prefix
        $ouMatch = $user.DistinguishedName -match 'OU=([^,]+)'
        $ou = if ($ouMatch) { $Matches[1] } else { "No OU" }
        $numericPrefix = if ($ou -match '^\d{2}$') { $Matches[0] } else { "999" }
        
        # Get and process each group
        $groups = $user.MemberOf | ForEach-Object {
            try {
                (Get-ADGroup $_).Name
            } catch {
                Write-Warning "Could not resolve group for user $($user.SamAccountName): $_"
                return "Unknown Group"
            }
        } | Sort-Object
        
        foreach ($group in $groups) {
            if (-not $groupColors.ContainsKey($group)) {
                $groupColors[$group] = $colorIndex
                $colorIndex++
                if ($colorIndex -gt 46) { $colorIndex = 35 }
            }
            
            $userData += [PSCustomObject]@{
                SortPrefix = $numericPrefix
                OU = $ou
                UserName = $user.Name
                SamAccountName = $user.SamAccountName
                Group = $group
                Comment = $user.Comment
                ColorIndex = $groupColors[$group]
            }
        }
    }
    
    return @{
        Data = $userData
        Colors = $groupColors
    }
}

function Export-ToCSV {
    param (
        [Parameter(Mandatory)]
        [object[]]$Data,
        [Parameter(Mandatory)]
        [string]$Path
    )
    
    $directory = Split-Path -Parent $Path
    if (-not (Test-Path $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }
    
    if (Test-Path $Path) {
        Remove-Item $Path -Force -ErrorAction Stop
    }
    
    $Data | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Force
    Write-Host "CSV export completed: $Path" -ForegroundColor Green
}

function Export-ToExcel {
    param (
        [Parameter(Mandatory)]
        [object[]]$Data,
        [Parameter(Mandatory)]
        [hashtable]$GroupColors,
        [Parameter(Mandatory)]
        [string]$Path
    )
    
    $excel = $null
    
    try {
        $directory = Split-Path -Parent $Path
        if (-not (Test-Path $directory)) {
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
        }
        
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        
        # Headers
        $headers = @("OU", "Benutzer", "SamAccountName", "Gruppe", "Kommentar")
        1..5 | ForEach-Object { 
            $worksheet.Cells.Item(1, $_) = $headers[$_ - 1]
        }
        
        # Data and colors
        $row = 2
        foreach ($item in $Data) {
            $worksheet.Cells.Item($row, 1) = $item.OU
            $worksheet.Cells.Item($row, 2) = $item.UserName
            $worksheet.Cells.Item($row, 3) = $item.SamAccountName
            $worksheet.Cells.Item($row, 4) = $item.Group
            $worksheet.Cells.Item($row, 5) = $item.Comment
            
            # Color only group cell
            $groupCell = $worksheet.Cells.Item($row, 4)
            $groupCell.Interior.ColorIndex = $GroupColors[$item.Group]
            
            $row++
        }
        
        # Format headers
        $headerRange = $worksheet.Range($worksheet.Cells(1, 1), $worksheet.Cells(1, 5))
        $headerRange.Font.Bold = $true
        $headerRange.Interior.ColorIndex = 15
        
        # Filters and column widths
        $worksheet.Range($worksheet.Cells(1, 1), $worksheet.Cells($row - 1, 5)).AutoFilter() | Out-Null
        $worksheet.Columns.Item(1).ColumnWidth = 20  # OU
        $worksheet.Columns.Item(2).ColumnWidth = 30  # Benutzer
        $worksheet.Columns.Item(3).ColumnWidth = 20  # SamAccountName
        $worksheet.Columns.Item(4).ColumnWidth = 50  # Gruppe
        $worksheet.Columns.Item(5).ColumnWidth = 40  # Kommentar
        
        if (Test-Path $Path) {
            Remove-Item $Path -Force
        }
        
        $workbook.SaveAs($Path)
        $workbook.Close($true)
        
        Write-Host "Excel export completed: $Path" -ForegroundColor Green
    }
    finally {
        if ($excel) {
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }
}

try {
    Write-Host "Starting AD user report generation..." -ForegroundColor Green
    
    # Force close Excel processes
    Close-ExcelProcesses
    
    # Get user data
    $result = Get-ADUserGroupData
    if (-not $result) { exit 0 }
    
    # Sort data with priority on 2-digit OUs
    $sortedData = $result.Data | Sort-Object {
        if ($_.SortPrefix -match '^\d{2}$') { 
            [int]$_.SortPrefix 
        } else { 
            999 
        }
    }, UserName, Group | Select-Object OU, UserName, SamAccountName, Group, Comment
    
    Write-Host "Processing $($sortedData.Count) user-group combinations..." -ForegroundColor Green
    
    # Export to CSV and Excel
    Export-ToCSV -Data $sortedData -Path $OutputPath
    Export-ToExcel -Data $sortedData -GroupColors $result.Colors -Path $ExcelPath
    
    Write-Host "Script completed successfully." -ForegroundColor Green
    Read-Host "Press Enter to exit"
}
catch {
    Write-Error "An error occurred: $_"
    Read-Host "Press Enter to exit"
    exit 1
}
