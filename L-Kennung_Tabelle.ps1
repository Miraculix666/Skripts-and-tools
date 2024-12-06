[CmdletBinding()]
param (
    [Parameter()]
    [string]$OutputPath = "C:\daten\AD_Benutzer_Gruppen_L.csv",
    [Parameter()]
    [string]$SearchPattern = "L*"
)

function Get-ADUserData {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$SearchPattern
    )
    
    try {
        $users = Get-ADUser -Filter "SamAccountName -like '$SearchPattern'" -Properties SamAccountName, Name, MemberOf, DistinguishedName
        
        $table = @()
        foreach ($user in $users) {
            Write-Verbose "Processing user '$($user.SamAccountName)'..."
            
            $ou = $user.DistinguishedName -replace '^CN=[^,]+,OU=([^,]+),.*$', '$1'
            
            if ($ou -eq "Benutzer") { continue }
            
            $groups = $user.MemberOf | Get-ADGroup | Select-Object -ExpandProperty Name | Sort-Object
            
            $table += [PSCustomObject]@{
                OU = $ou
                Benutzer = $user.Name
                Gruppen = $groups -join ", "
            }
        }
        
        return $table | Sort-Object -Property OU, Benutzer
    }
    catch {
        throw "Error retrieving AD user data: $_"
    }
}

function Export-ToCSV {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object[]]$Data,
        [Parameter(Mandatory)]
        [string]$Path
    )
    
    try {
        # Force close any open handles to the file
        if (Test-Path $Path) {
            Write-Verbose "Removing existing file '$Path'..."
            
            # Kill any Excel processes that might have the file open
            Get-Process | Where-Object { 
                $_.ProcessName -eq "EXCEL" -and $_.MainWindowTitle -like "*$Path*" 
            } | ForEach-Object {
                Write-Verbose "Closing Excel process that has the file open..."
                $_.CloseMainWindow()
                Start-Sleep -Seconds 1
                if (!$_.HasExited) {
                    $_.Kill()
                }
            }
            
            # Force delete the file
            Remove-Item $Path -Force -ErrorAction Stop
            Write-Verbose "Existing file deleted."
        }
        
        $Data | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
        Write-Verbose "Data exported to '$Path'"
        return $Path
    }
    catch {
        throw "Error exporting to CSV: $_"
    }
}

function Export-ToExcel {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$CsvPath
    )
    
    try {
        # Close any existing Excel instances with our file
        Get-Process | Where-Object { 
            $_.ProcessName -eq "EXCEL" -and $_.MainWindowTitle -like "*$CsvPath*" 
        } | ForEach-Object {
            Write-Verbose "Closing existing Excel instance..."
            $_.CloseMainWindow()
            Start-Sleep -Seconds 1
            if (!$_.HasExited) {
                $_.Kill()
            }
        }
        
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        
        $excel.DisplayAlerts = $false
        $csv = Get-Content $CsvPath | ConvertFrom-Csv
        $worksheet.Range("A1").CurrentRegion.Value2 = $csv
        $excel.DisplayAlerts = $true
        
        $headerRange = $worksheet.Range("A1:C1")
        $headerRange.Font.Bold = $true
        $headerRange.Interior.ColorIndex = 15
        
        $worksheet.AutoFilterMode = $false
        $headerRange.AutoFilter() | Out-Null
        
        $groups = $csv | Select-Object -ExpandProperty Gruppen -Unique
        $colors = 10..60
        
        foreach ($group in $groups) {
            $colorIndex = $colors[$groups.IndexOf($group) % $colors.Count]
            $cell = $worksheet.Cells.Find($group)
            
            while ($cell -ne $null) {
                $cell.EntireRow.Interior.ColorIndex = $colorIndex
                $cell = $worksheet.Cells.FindNext($cell)
            }
        }
        
        $worksheet.Columns.AutoFit()
        
        Write-Verbose "Excel report created successfully"
    }
    catch {
        throw "Error creating Excel report: $_"
    }
}

try {
    Write-Verbose "Starting AD user report generation..."
    
    Write-Verbose "Retrieving AD users starting with '$SearchPattern'..."
    $userData = Get-ADUserData -SearchPattern $SearchPattern
    
    Write-Verbose "Exporting data to CSV..."
    $csvPath = Export-ToCSV -Data $userData -Path $OutputPath
    
    Write-Verbose "Creating Excel report..."
    Export-ToExcel -CsvPath $csvPath
    
    Write-Verbose "Script completed successfully."
}
catch {
    Write-Error "An error occurred: $_"
    exit 1
}
