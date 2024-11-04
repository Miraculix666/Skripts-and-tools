# Liste der PCs und Nutzer
$pcs = "c26pbze70004600".."c26pbze70004613"
$users = "L1101061".."L1101061"
$password = "deinPasswort"  # Passwort für alle Nutzer

foreach ($i in 0..($pcs.Length - 1)) {
    $pc = $pcs[$i]
    $user = $users[$i]
    $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential ($user, $securePassword)

    # Öffne RDP-Sitzung
    Start-Process -FilePath "mstsc.exe" -ArgumentList "/v:$pc /prompt"

    # Warte kurz, um sicherzustellen, dass die Verbindung hergestellt ist
    Start-Sleep -Seconds 5

    # Starte die Anwendung auf dem entfernten PC
    Invoke-Command -ComputerName $pc -Credential $credential -ScriptBlock {
        Start-Process -FilePath "C:\Programme\x\x.exe"
    }
}
