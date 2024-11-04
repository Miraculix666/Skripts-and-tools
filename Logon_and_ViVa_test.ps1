# Liste der PCs und Nutzer als Strings
$pcs = @("c26pbze70004600", "c26pbze70004601", "c26pbze70004602", "c26pbze70004603", "c26pbze70004604", "c26pbze70004605", "c26pbze70004606", "c26pbze70004607", "c26pbze70004608", "c26pbze70004609", "c26pbze70004610", "c26pbze70004611", "c26pbze70004612", "c26pbze70004613")
$users = @("L1101061", "L1101062", "L1101063", "L1101064", "L1101065", "L1101066", "L1101067", "L1101068", "L1101069", "L1101070", "L1101071", "L1101072") 
$password = "PW"  # Passwort für alle Nutzer

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
