[CmdletBinding()]
param (
    [Parameter()]
    [string]$OutputPath = "C:\daten\AD_Benutzer_Gruppen_L.csv",
    [Parameter()]
    [string]$SearchPattern = "L*",
    [Parameter()]
    [string]$ExcelPath = "C:\daten\AD_Benutzer_Gruppen_L.xlsx"
)

try {
    Write-Host "Starting AD user report generation..." -ForegroundColor Green
    
    # Get AD users and process data
    Write-Verbose "Retrieving AD users matching pattern '$SearchPattern'..."
    $users = Get-ADUser -Filter "SamAccountName -like '$SearchPattern'" -Properties SamAccountName, Name, MemberOf, DistinguishedName
    
    if ($users.Count -eq 0) {
        Write-Warning "No users found matching pattern '$SearchPattern'"
        exit 0
    }
    
    # Process user data
    $table = @()
    foreach ($user in $users) {
        Write-Verbose "Processing user '$($user.SamAccountName)'..."
        
        # Extract full OU path and numeric prefix
        $ouMatch = $user.DistinguishedName -match 'OU=([^,]+)'
        $ou = if ($ouMatch) { $Matches[1] } else { "No OU" }
        $numericPrefix = if ($ou -match '^\d{2,3}') { $Matches[0] } else { "999" }
        
        # Get group memberships
        $groups = $user.MemberOf | ForEach-Object {
            try {
                Get-ADGroup $_ | Select-Object -ExpandProperty Name
            } catch {
                Write-Warning "Could not resolve group $_"
                return "Unknown Group"
            }
        } | Sort-Object
        
        # Create custom object with numeric prefix for sorting
        $table += [PSCustomObject]@{
            NumericPrefix = $numericPrefix
            OU = $ou
            Benutzer = $user.Name
            SamAccountName = $user.SamAccountName
            Gruppen = $groups -join "; "
            GroupCount = ($groups | Measure-Object).Count
        }
    }
    
    # Sort data
    $processedData = $table | Sort-Object -Property NumericPrefix, OU, Benutzer | 
                             Select-Object OU, Benutzer, SamAccountName, Gruppen, GroupCount
    
    Write-Host "Found $($processedData.Count) users" -ForegroundColor Green
    
    # Export to CSV
    try {
        $directory = Split-Path -Parent $OutputPath
        if (-not (Test-Path $directory)) {
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
        }
        
        if (Test-Path $OutputPath) {
            Stop-Process -Name "excel" -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 1
            Remove-Item $OutputPath -Force -ErrorAction Stop
        }
        
        $processedData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -Force
        Write-Host "CSV export completed: $OutputPath" -ForegroundColor Green
    }
    catch {
        throw "Error exporting to CSV: $_"
    }
    
    # Export to Excel
    try {
        $directory = Split-Path -Parent $ExcelPath
        if (-not (Test-Path $directory)) {
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
        }
        
        Stop-Process -Name "excel" -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
        
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        
        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        
        # Add headers
        $headers = @("OU", "Benutzer", "SamAccountName", "Gruppen", "Anzahl Gruppen")
        for ($i = 0; $i -lt $headers.Count; $i++) {
            $worksheet.Cells.Item(1, $i + 1) = $headers[$i]
        }
        
        # Add data
        $row = 2
        foreach ($item in $processedData) {
            $worksheet.Cells.Item($row, 1) = $item.OU
            $worksheet.Cells.Item($row, 2) = $item.Benutzer
            $worksheet.Cells.Item($row, 3) = $item.SamAccountName
            $worksheet.Cells.Item($row, 4) = $item.Gruppen
            $worksheet.Cells.Item($row, 5) = $item.GroupCount
            $row++
        }
        
        # Format headers
        $headerRange = $worksheet.Range($worksheet.Cells(1,1), $worksheet.Cells(1,5))
        $headerRange.Font.Bold = $true
        $headerRange.Interior.ColorIndex = 15
        
        # Add filters
        $worksheet.Range($worksheet.Cells(1,1), $worksheet.Cells($row-1,5)).AutoFilter() | Out-Null
        
        # Format columns
        $worksheet.Columns.Item(1).ColumnWidth = 20  # OU
        $worksheet.Columns.Item(2).ColumnWidth = 30  # Benutzer
        $worksheet.Columns.Item(3).ColumnWidth = 20  # SamAccountName
        $worksheet.Columns.Item(4).ColumnWidth = 50  # Gruppen
        $worksheet.Columns.Item(5).ColumnWidth = 15  # Anzahl Gruppen
        
        # Color alternate rows
        $dataRange = $worksheet.Range($worksheet.Cells(2,1), $worksheet.Cells($row-1,5))
        $dataRange.Interior.ColorIndex = 2
        
        # Save and close
        $workbook.SaveAs($ExcelPath)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        
        Write-Host "Excel export completed: $ExcelPath" -ForegroundColor Green
    }
    catch {
        throw "Error creating Excel report: $_"
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
