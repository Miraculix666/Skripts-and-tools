# Module for exporting reports
function Export-ReportsEnhanced {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object[]]$DetailedData,
        [Parameter(Mandatory)]
        [object[]]$MatrixData,
        [Parameter(Mandatory)]
        [hashtable]$Groups,
        [Parameter(Mandatory)]
        [string]$CsvPath,
        [Parameter(Mandatory)]
        [string]$ExcelPath
    )

    try {
        Export-ToCsvEnhanced -Data $DetailedData -Path $CsvPath
        Export-ToExcelEnhanced -DetailedData $DetailedData -MatrixData $MatrixData -Groups $Groups -Path $ExcelPath
    }
    catch {
        throw "Error exporting reports: $_"
    }
}

function Export-ToCsvEnhanced {
    [CmdletBinding()]
    param ($Data, $Path)

    try {
        New-OutputDirectory -Path $Path
        Remove-ExistingFile -Path $Path

        $Data | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Force
        Write-Host "CSV export completed: $Path" -ForegroundColor Green
    }
    catch {
        throw "Error exporting to CSV: $_"
    }
}

function Export-ToExcelEnhanced {
    [CmdletBinding()]
    param ($DetailedData, $MatrixData, $Groups, $Path)

    $excel = $null

    try {
        New-OutputDirectory -Path $Path
        Close-ExcelProcesses
        Remove-ExistingFile -Path $Path

        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Add()

        # Detailed View Sheet
        $detailedSheet = $workbook.Worksheets.Item(1)
        $detailedSheet.Name = "Detailed View"
        Set-ExcelDetailedSheetProperties -Worksheet $detailedSheet -Data $DetailedData

        # Matrix View Sheet
        $matrixSheet = $workbook.Worksheets.Add()
        $matrixSheet.Name = "Matrix View"
        Set-ExcelMatrixSheetProperties -Worksheet $matrixSheet -Data $MatrixData -Groups $Groups

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

function Set-ExcelDetailedSheetProperties {
    [CmdletBinding()]
    param (
        $Worksheet,
        $Data
    )

    # Add headers
    $headers = @("OU", "Benutzer", "SamAccountName", "Gruppe", "Kommentar")
    1..5 | ForEach-Object {
        $Worksheet.Cells.Item(1, $_) = $headers[$_ - 1]
    }

    # Add data and format cells
    $row = 2
    foreach ($item in $Data) {
        $Worksheet.Cells.Item($row, 1) = $item.OU
        $Worksheet.Cells.Item($row, 2) = $item.Benutzer
        $Worksheet.Cells.Item($row, 3) = $item.SamAccountName
        $Worksheet.Cells.Item($row, 4) = $item.Gruppe
        $Worksheet.Cells.Item($row, 5) = $item.Kommentar

        # Color only the group cell
        $groupCell = $Worksheet.Cells.Item($row, 4)
        $groupCell.Interior.ColorIndex = $item.ColorIndex

        $row++
    }

    # Format headers
    $headerRange = $Worksheet.Range($Worksheet.Cells(1, 1), $Worksheet.Cells(1, 5))
    $headerRange.Font.Bold = $true
    $headerRange.Interior.ColorIndex = 15

    # Add filters and adjust column widths
    $Worksheet.Range($Worksheet.Cells(1, 1), $Worksheet.Cells($row - 1, 5)).AutoFilter() | Out-Null
    $Worksheet.Columns.Item(1).ColumnWidth = 40  # OU
    $Worksheet.Columns.Item(2).ColumnWidth = 30  # Benutzer
    $Worksheet.Columns.Item(3).ColumnWidth = 20  # SamAccountName
    $Worksheet.Columns.Item(4).ColumnWidth = 40  # Gruppe
    $Worksheet.Columns.Item(5).ColumnWidth = 50  # Kommentar
}

function Set-ExcelMatrixSheetProperties {
    [CmdletBinding()]
    param (
        $Worksheet,
        $Data,
        $Groups
    )

    # Add headers
    $sortedGroups = $Groups.Keys | Sort-Object
    $matrixHeaders = @("OU", "Benutzer", "SamAccountName") + $sortedGroups
    1..$matrixHeaders.Count | ForEach-Object {
        $Worksheet.Cells.Item(1, $_) = $matrixHeaders[$_ - 1]
    }

    # Add data
    $row = 2
    foreach ($item in $Data) {
        $Worksheet.Cells.Item($row, 1) = $item.OUPath
        $Worksheet.Cells.Item($row, 2) = $item.Benutzer
        $Worksheet.Cells.Item($row, 3) = $item.SamAccountName

        $col = 4
        foreach ($group in $sortedGroups) {
            if ($item.$group -eq "X") {
                $cell = $Worksheet.Cells.Item($row, $col)
                $cell.Value = "X"
                $cell.Interior.ColorIndex = $Groups[$group]
            }
            $col++
        }
        $row++
    }

    # Format headers
    $headerRange = $Worksheet.Range($Worksheet.Cells(1, 1), $Worksheet.Cells(1, $matrixHeaders.Count))
    $headerRange.Font.Bold = $true
    $headerRange.Interior.ColorIndex = 15

    # Add filters and adjust column widths
    $Worksheet.Range($Worksheet.Cells(1, 1), $Worksheet.Cells($row - 1, $matrixHeaders.Count)).AutoFilter() | Out-Null
    $Worksheet.Columns.Item(1).ColumnWidth = 40  # OU
    $Worksheet.Columns.Item(2).ColumnWidth = 30  # Benutzer
    $Worksheet.Columns.Item(3).ColumnWidth = 20  # SamAccountName
    4..$matrixHeaders.Count | ForEach-Object {
        $Worksheet.Columns.Item($_).ColumnWidth = 15  # Group columns
    }
}

Export-ModuleMember -Function Export-ReportsEnhanced
