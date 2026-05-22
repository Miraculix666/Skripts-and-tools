$testSize = 10000
$allUsers = 1..$testSize | ForEach-Object {
    [PSCustomObject]@{
        SamAccountName = "User$_"
    }
}

$testFile1 = "test1.txt"
$testFile2 = "test2.txt"

$timeBefore = Measure-Command {
    $allUsers | ForEach-Object { $_.SamAccountName | Out-File -Append -FilePath $testFile1 }
}

$timeAfter = Measure-Command {
    $allUsers | Select-Object -ExpandProperty SamAccountName | Out-File -Append -FilePath $testFile2
}

Write-Host "Time before (Foreach-Object { Out-File }): $($timeBefore.TotalMilliseconds) ms"
Write-Host "Time after (Out-File once): $($timeAfter.TotalMilliseconds) ms"
