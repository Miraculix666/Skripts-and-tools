# Liste der PCs und Nutzer als Strings
$pcs = @("c26pbze70004600", "c26pbze70004601", "c26pbze70004602", "c26pbze70004603", "c26pbze70004604", "c26pbze70004605", "c26pbze70004606", "c26pbze70004607", "c26pbze70004608", "c26pbze70004609", "c26pbze70004610", "c26pbze70004611", "c26pbze70004612", "c26pbze70004613")
$users = @("L1101061", "L1101062", "L1101063", "L1101064", "L1101065", "L1101066", "L1101067", "L1101068", "L1101069", "L1101070", "L1101071", "L1101072") 

# Frage nach dem Passwort für RDP und Benutzer
$rdpPassword = Read-Host -AsSecureString "Bitte gib das RDP-Passwort ein"
$userPassword = Read-Host -AsSecureString "Bitte gib das Passwort für die anderen Benutzer ein"

foreach ($i in 0..($pcs.Length - 1)) {
    $pc = $pcs[$i]
    $user = $users[$i]
    $securePassword = $userPassword
    $credential = New-Object System.Management.Automation.PSCredential ($user, $securePassword)

    # RDP-Verbindung im Speicher vorbereiten
    $encodedRdpPassword = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes((New-Object System.Management.Automation.PSCredential ("user1", $rdpPassword)).GetNetworkCredential().Password))
    $rdpArgs = @"
screen mode id:i:1
desktopwidth:i:1024
desktopheight:i:768
session bpp:i:32
winposstr:s:0,1,10,10,800,600
full address:s:$pc
compression:i:1
keyboardhook:i:2
audiomode:i:0
redirectprinters:i:1
redirectcomports:i:0
redirectsmartcards:i:1
redirectclipboard:i:1
redirectposdevices:i:0
drivestoredirect:s:
autoreconnection enabled:i:1
password 51:b:$encodedRdpPassword
"@

    # Öffne RDP-Sitzung mit dem aktuellen Benutzer
    Start-Process -FilePath "mstsc.exe" -ArgumentList $rdpArgs

    # Warte kurz, um sicherzustellen, dass die Verbindung hergestellt ist
    Start-Sleep -Seconds 5

    # Starte die Anwendung auf dem entfernten PC im Kontext eines anderen Benutzers
    Invoke-Command -ComputerName $pc -Credential $credential -ScriptBlock {
        Start-Process -FilePath "C:\Programme\x\x.exe"
    }
}
