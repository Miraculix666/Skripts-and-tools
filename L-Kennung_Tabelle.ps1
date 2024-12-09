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
            if (!$_.HasExited) {
                $_.Kill()
            }
        } catch {
            Write-Warning "Could not close Excel process: $_"
        }
    }
    Start-Sleep -Seconds 2
}

try {
    Write-Host "Starting AD user report generation..." -ForegroundColor Green
    
    # Get all users with SamAccountName starting with 'L'
    $users = Get-ADUser -Filter "SamAccountName -like 'L*'" -Properties SamAccountName, Name, MemberOf, DistinguishedName
    
    if (-not $users) {
        Write-Warning "No users found with SamAccountName starting with 'L'"
        exit 0
    }
    
    # Initialize collections
    $userData = @()
    $groupColors = @{}
    $colorIndex = 20  # Excel color index starting point
    
    foreach ($user in $users) {
        Write-Verbose "Processing user: $($user.SamAccountName)"
        
        # Extract OU and numeric prefix
        $ouMatch = $user.DistinguishedName -match 'OU=([^,]+)'
        $ou = if ($ouMatch) { $Matches[1] } else { "No OU" }
        $numericPrefix = if ($ou -match '^\d{2,3}') { $Matches[0] } else { "999" }
        
        # Get and process each group
        $groups = $user.MemberOf | ForEach-Object {
            try {
                (Get-ADGroup $_).Name
            } catch {
                Write-Warning "Could not resolve group for user $($user.SamAccountName): $_"
                return "Unknown Group"
            }
        } | Sort-Object
        
        # Create entry for each group
        foreach ($group in $groups) {
            # Assign color if not already assigned
            if (-not $groupColors.ContainsKey($group)) {
                $groupColors[$group] = $colorIndex
                $colorIndex++
                if ($colorIndex -gt 56) { $colorIndex = 20 }  # Reset if we run out of colors
            }
            
            $userData += [PSCustomObject]@{
                SortPrefix = $numericPrefix
                OU = $ou
                UserName = $user.Name
                SamAccountName = $user.SamAccountName
                Group = $group
                ColorIndex = $groupColors[$group]
            }
        }
    }
    
    # Sort data
    $sortedData = $userData | Sort-Object SortPrefix, OU, UserName, Group |
                            Select-Object OU, UserName, SamAccountName, Group
    
    Write-Host "Processing $($sortedData.Count) user-group combinations..." -ForegroundColor Green
    
    # Export to CSV
    try {
        $csvDir = Split-Path -Parent $OutputPath
        if (-not (Test-Path $csvDir)) {
            New-Item -ItemType Directory -Path $csvDir -Force | Out-Null
        }
        
        # Force close any open files
        if (Test-Path $OutputPath) {
            Remove-Item $OutputPath -Force -ErrorAction Stop
        }
        
        $sortedData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -Force
        Write-Host "CSV export completed: $OutputPath" -ForegroundColor Green
    }
    catch {
        throw "Failed to export CSV: $_"
    }
    
    # Export to Excel
    try {
        $excelDir = Split-Path -Parent $ExcelPath
        if (-not (Test-Path $excelDir)) {
            New-Item -ItemType Directory -Path $excelDir -Force | Out-Null
        }
        
        # Close any open Excel processes
        Close-ExcelProcesses
        
        # Create new Excel file
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        
        # Add headers
        $headers = @("OU", "Benutzer", "SamAccountName", "Gruppe")
        1..4 | ForEach-Object { 
            $worksheet.Cells.Item(1, $_) = $headers[$_ - 1]
        }
        
        # Add data with colors
        $row = 2
        foreach ($item in $sortedData) {
            $worksheet.Cells.Item($row, 1) = $item.OU
            $worksheet.Cells.Item($row, 2) = $item.UserName
            $worksheet.Cells.Item($row, 3) = $item.SamAccountName
            $worksheet.Cells.Item($row, 4) = $item.Group
            
            # Color the row
            $range = $worksheet.Range($worksheet.Cells($row, 1), $worksheet.Cells($row, 4))
            $range.Interior.ColorIndex = $groupColors[$item.Group]
            
            $row++
        }
        
        # Format headers
        $headerRange = $worksheet.Range($worksheet.Cells(1, 1), $worksheet.Cells(1, 4))
        $headerRange.Font.Bold = $true
        $headerRange.Interior.ColorIndex = 15
        
        # Add filters and adjust column widths
        $worksheet.Range($worksheet.Cells(1, 1), $worksheet.Cells($row - 1, 4)).AutoFilter() | Out-Null
        $worksheet.Columns.Item(1).ColumnWidth = 20  # OU
        $worksheet.Columns.Item(2).ColumnWidth = 30  # Benutzer
        $worksheet.Columns.Item(3).ColumnWidth = 20  # SamAccountName
        $worksheet.Columns.Item(4).ColumnWidth = 50  # Gruppe
        
        # Save and close
        if (Test-Path $ExcelPath) {
            Remove-Item $ExcelPath -Force
        }
        
        $workbook.SaveAs($ExcelPath)
        $workbook.Close($true)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        Write-Host "Excel export completed: $ExcelPath" -ForegroundColor Green
    }
    catch {
        throw "Failed to export Excel file: $_"
    }
    finally {
        if ($excel) {
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
    }
    
    Write-Host "Script completed successfully." -ForegroundColor Green
    Read-Host "Press Enter to exit"
}
catch {
    Write-Error "An error occurred: $_"
    Read-Host "Press Enter to exit"
    exit 1
}
