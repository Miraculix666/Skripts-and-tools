# Liste der PCs und Nutzer als Strings
$pcs = @("c00", "c01", "c02", "c03", "c04", "c05", "c06", "c07", "c08", "c09", "c10", "c11", "c12", "c13")
$users = @("N00", "N01", "N02", "N03", "N04", "N05", "N06", "N07", "N08", "N09", "N10", "N11", "N12", "N13")
$password = "deinPasswort"  # Passwort für alle Nutzer

foreach ($i in 0..($pcs.Length - 1)) {
    $pc = $pcs[$i]
    $user = $users[$i]
    $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential ($user, $securePassword)

    # Öffne RDP-Sitzung mit dem aktuellen Benutzer
    Start-Process -FilePath "mstsc.exe" -ArgumentList "/v:$pc /prompt"

    # Warte kurz, um sicherzustellen, dass die Verbindung hergestellt ist
    Start-Sleep -Seconds 5

    # Starte die Anwendung auf dem entfernten PC im Kontext eines anderen Benutzers
    Invoke-Command -ComputerName $pc -Credential $credential -ScriptBlock {
        Start-Process -FilePath "C:\Programme\x\x.exe"
    }
}
