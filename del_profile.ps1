### scrit too delete remotly user rofiles on a comuter

param (
    [Parameter(Mandatory=$true)]
    [string]$ComputerName,

    [Parameter(Mandatory=$true)]
    [string]$UserName
)

Invoke-Command -ComputerName $ComputerName -ScriptBlock {
    param($user)
    # Benutzer abmelden
    query user /server:$env:computername | Where-Object {$_ -match $user} | ForEach-Object {
        logoff ($_ -split '\s+')[2] /server:$env:computername
    }
    # Warten, bis der Benutzer abgemeldet ist
    Start-Sleep -Seconds 10
    # Benutzerprofil löschen
    $localpath = 'C:\Users\' + $user
    Get-WmiObject -Class Win32_UserProfile | Where-Object {$_.LocalPath -eq $localpath} | Remove-WmiObject
} -ArgumentList $UserName
