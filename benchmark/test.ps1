$Computers = 1..1000 | ForEach-Object { "Computer$_" }

Write-Host "Baseline array addition (+=)..."
$startTime = Get-Date
$jobs1 = @()
foreach ($computer in $Computers) {
    $job = [PSCustomObject]@{
        ComputerName = $computer
        Status = "Pending"
    }
    $jobs1 += $job
}
$endTime = Get-Date
$time1 = ($endTime - $startTime).TotalMilliseconds
Write-Host "Baseline took: $time1 ms"

Write-Host "Optimized assignment pipeline..."
$startTime = Get-Date
$jobs2 = foreach ($computer in $Computers) {
    [PSCustomObject]@{
        ComputerName = $computer
        Status = "Pending"
    }
}
$endTime = Get-Date
$time2 = ($endTime - $startTime).TotalMilliseconds
Write-Host "Optimized took: $time2 ms"
Write-Host "Improvement: $([math]::Round($time1 / $time2, 2))x"
