# Importiere das Configuration Manager PowerShell-Modul
Import-Module 'C:\Program Files (x86)\Microsoft Endpoint Manager\AdminConsole\bin\ConfigurationManager.psd1'

# Verbinde dich mit dem Configuration Manager
$SiteCode = "ABC"  # Ersetze "ABC" durch deinen Site-Code
cd "$SiteCode:"

# Hole alle Anwendungen und ihren Installationsstatus
$applications = Get-CMApplication | Select-Object -Property LocalizedDisplayName, SoftwareVersion, IsDeployed

# Ausgabe der Anwendungen und ihres Status
$applications | Format-Table -Property LocalizedDisplayName, SoftwareVersion, IsDeployed -AutoSize
