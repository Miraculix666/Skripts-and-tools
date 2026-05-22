# Main script for AD User Groups Report
[CmdletBinding()]
param (
    [Parameter()]
    [string]$SearchPattern = "L*",

    [Parameter()]
    [string]$BaseDN = "DC=domain,DC=local",

    [Parameter()]
    [string[]]$ExcludeOUs = @("Benutzer", "Disabled Users"),

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

    # Get and process AD user data
    $result = Get-ADUserDataEnhanced -SearchPattern $SearchPattern -SearchBase $BaseDN -ExcludeOUs $ExcludeOUs
    if (-not $result) { exit 0 }

    # Sort data for Detailed View
    $sortedDetailedData = $result.DetailedData | Sort-Object {
        if ($_.SortPrefix -match '^\d{2,3}$') { [int]$_.SortPrefix } else { 999 }
    }, Benutzer, Gruppe

    # Sort data for Matrix View
    $sortedMatrixData = $result.MatrixData | Sort-Object {
        if ($_.NumericPrefix -match '^\d{2,3}$') { [int]$_.NumericPrefix } else { 999 }
    }, Benutzer

    Write-Host "Processing $($sortedDetailedData.Count) user-group combinations..." -ForegroundColor Green

    # Export data in all formats
    Export-ReportsEnhanced -DetailedData $sortedDetailedData -MatrixData $sortedMatrixData -Groups $result.Groups -CsvPath $OutputPath -ExcelPath $ExcelPath

    # Export visualizations
    if (Get-Command Export-VisualizationsEnhanced -ErrorAction SilentlyContinue) {
        Export-VisualizationsEnhanced -Data $sortedDetailedData -HtmlPath $HtmlPath -VisioPath $VisioPath -MindManagerPath $MindManagerPath
    }

    Write-Host "Script completed successfully." -ForegroundColor Green
    Read-Host "Press Enter to exit"
}
catch {
    Write-Error "An error occurred: $_"
    Read-Host "Press Enter to exit"
    exit 1
}
