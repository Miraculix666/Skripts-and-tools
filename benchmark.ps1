$iterations = 1000

Write-Host "Benchmarking @() +="
$time1 = Measure-Command {
    $arr = @()
    for ($i = 0; $i -lt $iterations; $i++) {
        $arr += $i
    }
}
Write-Host "Time: $($time1.TotalMilliseconds) ms"

Write-Host "Benchmarking [System.Collections.Generic.List[PSObject]]"
$time2 = Measure-Command {
    $list = [System.Collections.Generic.List[PSObject]]::new()
    for ($i = 0; $i -lt $iterations; $i++) {
        $list.Add($i)
    }
}
Write-Host "Time: $($time2.TotalMilliseconds) ms"
