$items = 1..10000

Write-Host "Testing += array addition:"
$time1 = Measure-Command {
    $myArray = @()
    foreach ($item in $items) {
        $myArray += $item
    }
}
Write-Host "Time: $($time1.TotalMilliseconds) ms"

Write-Host "Testing List[T] Add():"
$time2 = Measure-Command {
    $myList = New-Object System.Collections.Generic.List[int]
    foreach ($item in $items) {
        $myList.Add($item)
    }
}
Write-Host "Time: $($time2.TotalMilliseconds) ms"

Write-Host "Testing pipeline array assignment:"
$time3 = Measure-Command {
    $myArray = foreach ($item in $items) {
        $item
    }
}
Write-Host "Time: $($time3.TotalMilliseconds) ms"

Write-Host "Testing pipeline array assignment with @():"
$time4 = Measure-Command {
    $myArray = @(foreach ($item in $items) {
        $item
    })
}
Write-Host "Time: $($time4.TotalMilliseconds) ms"
