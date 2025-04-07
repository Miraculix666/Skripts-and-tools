# Erforderliche Module laden
Import-Module ActiveDirectory

# Dynamische OU-Suche
$domain = Get-ADDomain
$targetOUs = @('81', '82') | ForEach-Object {
    Get-ADOrganizationalUnit -Filter "Name -eq '$_'" -SearchBase $domain.DistinguishedName -SearchScope Subtree
}

if (-not $targetOUs) {
    Write-Host "Fehler: OUs 81/82 nicht gefunden!" -ForegroundColor Red
    exit
}

# Benutzer mit Namensfiltern finden
$users = $targetOUs | ForEach-Object {
    Get-ADUser -Filter "Name -like 'L11[04]*'" -SearchBase $_.DistinguishedName -Properties Enabled
}

if (-not $users) {
    Write-Host "Keine passenden Benutzer gefunden." -ForegroundColor Yellow
    exit
}

# Passwortabfrage
$plainPassword = Read-Host "Neues Passwort eingeben" -AsSecureString
$password = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($plainPassword)
)

# Initialisierung Zähler
$totalUsers = $users.Count
$successCount = 0
$errorUsers = @()

# Verarbeitungsschleife
foreach ($user in $users) {
    try {
        # Net User-Kommando für Basisoperationen
        $output = net user $user.SamAccountName $password /DOMAIN /ACTIVE:YES /PASSWORDCHG:NO 2>&1
        
        if ($LASTEXITCODE -ne 0) {
            throw "Net User Fehler: $($output -join ' ')"
        }

        # Zusätzliche Einstellungen mit AD-Modul
        Set-ADUser -Identity $user -PasswordNeverExpires $true -Replace @{
            lastLogonTimestamp = [DateTime]::Now.ToFileTime()
        } -ErrorAction Stop

        $successCount++
        Write-Host "Erfolg: $($user.SamAccountName)" -ForegroundColor Green
    }
    catch {
        $errorUsers += [PSCustomObject]@{
            Benutzer = $user.SamAccountName
            Fehler = $_.Exception.Message
        }
        Write-Host "Fehler: $($user.SamAccountName) - $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Zusammenfassung
Write-Host "`nZusammenfassung:"
Write-Host "Verarbeitete Benutzer: $totalUsers"
Write-Host "Erfolgreich: $successCount"
Write-Host "Fehler: $($errorUsers.Count)"

if ($errorUsers.Count -gt 0) {
    Write-Host "`nDetailfehler:"
    $errorUsers | Format-Table -AutoSize
}

# Aufräumen
Remove-Variable password, plainPassword -ErrorAction SilentlyContinue
