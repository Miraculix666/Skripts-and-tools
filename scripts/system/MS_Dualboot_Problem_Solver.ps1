# Erfordert Administrator-Rechte
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))  
{  
    Write-Warning "Bitte führen Sie dieses Skript als Administrator aus!"
    Break
}

# Secure Boot deaktivieren
Write-Host "Deaktiviere Secure Boot..."
Confirm-SecureBootUEFI -Disable

# BCD-Store sichern
bcdedit /export C:\BCD_Backup

# Linux-Bootloader-Eintrag hinzufügen (Beispiel für Ubuntu)
$linuxPath = "\EFI\ubuntu\grubx64.efi"
bcdedit /create /d "Ubuntu" /application osloader
$guid = (bcdedit /create /d "Ubuntu" /application osloader) -replace ".*({.*}).*", '$1'
bcdedit /set $guid device partition=C:
bcdedit /set $guid path $linuxPath
bcdedit /set $guid description "Ubuntu Linux"

# Bootmenü-Timeout einstellen (in Sekunden)
bcdedit /timeout 10

# Signaturen großer Linux-Distributionen eintragen
$certPath = "C:\LinuxCerts"
New-Item -Path $certPath -ItemType Directory -Force

# Ubuntu
$ubuntuCert = "https://raw.githubusercontent.com/rhboot/shim/main/ubuntu-ca.crt"
Invoke-WebRequest -Uri $ubuntuCert -OutFile "$certPath\ubuntu-ca.crt"

# Fedora
$fedoraCert = "https://getfedora.org/static/fedora.gpg"
Invoke-WebRequest -Uri $fedoraCert -OutFile "$certPath\fedora.gpg"

# OpenSUSE
$opensuseCert = "https://build.opensuse.org/projects/openSUSE:Factory/public_key"
Invoke-WebRequest -Uri $opensuseCert -OutFile "$certPath\opensuse.pub"

# Zertifikate importieren
Import-Certificate -FilePath "$certPath\ubuntu-ca.crt" -CertStoreLocation Cert:\LocalMachine\Root
Import-Certificate -FilePath "$certPath\fedora.gpg" -CertStoreLocation Cert:\LocalMachine\Root
Import-Certificate -FilePath "$certPath\opensuse.pub" -CertStoreLocation Cert:\LocalMachine\Root

Write-Host "Konfiguration abgeschlossen. Bitte starten Sie Ihren Computer neu, um die Änderungen zu übernehmen."
