[CmdletBinding()]
param (
    [Parameter()]
    [string]$OutputPath = "C:\daten\AD_Benutzer_Gruppen_L.csv",
    [Parameter()]
    [string]$ExcelPath = "C:\daten\AD_Benutzer_Gruppen_L.xlsx",
    [Parameter()]
    [string]$HtmlPath = "C:\daten\AD_Benutzer_Gruppen_L.html",
    [Parameter()]
    [string]$VisioPath = "C:\daten\AD_Benutzer_Gruppen_L.vsdx",
    [Parameter()]
    [string]$MindManagerPath = "C:\daten\AD_Benutzer_Gruppen_L.mmap"
)

# Import required modules
Import-Module "$PSScriptRoot\..\Modules\Utils.psm1" -Force
Import-Module "$PSScriptRoot\..\Modules\Get-ADUserData.psm1" -Force
Import-Module "$PSScriptRoot\..\Modules\Export-Reports.psm1" -Force
Import-Module "$PSScriptRoot\..\Modules\Export-Visualizations.psm1" -Force

try {
    Write-Host "Starting enhanced AD user report generation..." -ForegroundColor Green
    
    # Close any open Excel processes
    Close-ExcelProcesses
    
    # Get user data
    $result = Get-ADUserDataEnhanced -SearchPattern "L*"
    if (-not $result) {
        Write-Warning "No user data retrieved."
        exit 0
    }
    
    # Export reports
    Export-ReportsEnhanced -DetailedData $result.DetailedData -MatrixData $result.MatrixData -Groups $result.Groups -CsvPath $OutputPath -ExcelPath $ExcelPath
    
    # Generate visualizations
    Export-VisualizationsEnhanced -Data $result.DetailedData -HtmlPath $HtmlPath -VisioPath $VisioPath -MindManagerPath $MindManagerPath
    
    Write-Host "Script completed successfully." -ForegroundColor Green
    Read-Host "Press Enter to exit"
}
catch {
    Write-Error "An error occurred: $_"
    Read-Host "Press Enter to exit"
    exit 1
}