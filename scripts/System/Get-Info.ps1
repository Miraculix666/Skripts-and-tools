# --- 1. DATEN-SAMMLUNG (Der Deep Dive) ---
Write-Host "Sammle Informationen... Bitte warten." -ForegroundColor Cyan

# Basis Info
$domain = Get-ADDomain
$forest = Get-ADForest
$rootDSE = Get-ADRootDSE
$localComp = Get-CimInstance Win32_OperatingSystem
$cpu = Get-CimInstance Win32_Processor
$cs = Get-CimInstance Win32_ComputerSystem

# AD Details
$sites = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().Sites
$subnets = Get-ADReplicationSubnet -Filter *
$trusts = Get-ADTrust -Filter *
$schemaVersion = (Get-ADObject $rootDSE.schemaNamingContext -Property objectVersion).objectVersion

# Security Check (Counts)
$adminCount = (Get-ADGroupMember -Identity "Domain Admins").Count
$enterpriseAdmins = (Get-ADGroupMember -Identity "Enterprise Admins" -ErrorAction SilentlyContinue).Count
$inactiveUsers = (Search-ADAccount -AccountInactive -TimeSpan 90.00:00:00 -UsersOnly).Count

# Domain Controller Details
$dcs = Get-ADDomainController -Filter * | Select-Object Name, IPv4Address, Site, OperatingSystem, IsGlobalCatalog, IsReadOnly

# --- 2. HTML GENERIERUNG ---
$reportPath = "$env:TEMP\AD_DeepDive_Report.html"

# CSS & Header
$html = @"
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <title>AD Deep Dive Inventory</title>
    <style>
        :root { --primary: #0078d4; --bg: #f3f5f7; --card: #ffffff; --text: #201f1e; }
        body { font-family: 'Segoe UI', Tahoma, sans-serif; background: var(--bg); color: var(--text); padding: 20px; line-height: 1.6; }
        .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(450px, 1fr)); gap: 20px; }
        .card { background: var(--card); padding: 20px; border-radius: 4px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); border-top: 4px solid var(--primary); }
        h1 { color: var(--primary); margin-bottom: 30px; text-align: center; font-weight: 300; font-size: 2.5em; }
        h2 { border-bottom: 1px solid #eee; padding-bottom: 10px; color: var(--primary); font-size: 1.2em; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 0.9em; }
        th, td { text-align: left; padding: 8px; border-bottom: 1px solid #f0f0f0; }
        th { color: #666; font-weight: 600; width: 40%; }
        .badge { padding: 2px 8px; border-radius: 10px; font-size: 0.8em; background: #eee; }
        .critical { color: #d83b01; font-weight: bold; }
    </style>
</head>
<body>
    <h1>System & Directory Inventory Report</h1>
    
    <div class="grid">
        <div class="card">
            <h2>💻 Local Host & OS Info</h2>
            <table>
                <tr><th>FQDN</th><td>$($env:COMPUTERNAME).$($env:USERDNSDOMAIN)</td></tr>
                <tr><th>Operating System</th><td>$($localComp.Caption) (Build $($localComp.BuildNumber))</td></tr>
                <tr><th>Hardware</th><td>$($cs.Manufacturer) $($cs.Model)</td></tr>
                <tr><th>CPU / RAM</th><td>$($cpu.Name) / $([math]::Round($cs.TotalPhysicalMemory / 1GB, 0)) GB</td></tr>
                <tr><th>Last Boot</th><td>$($localComp.LastBootUpTime)</td></tr>
            </table>
        </div>

        <div class="card">
            <h2>🌳 AD Forest & Domain</h2>
            <table>
                <tr><th>Forest Root</th><td>$($forest.RootDomain)</td></tr>
                <tr><th>Domain FQDN</th><td>$($domain.DNSRoot)</td></tr>
                <tr><th>NetBIOS / SID</th><td>$($domain.NetBIOSName) / $($domain.DomainSID)</td></tr>
                <tr><th>Functional Level</th><td>F: $($forest.ForestMode) / D: $($domain.DomainMode)</td></tr>
                <tr><th>Schema Version</th><td>$schemaVersion (AD 2019/2022)</td></tr>
            </table>
        </div>

        <div class="card">
            <h2>📍 Sites & Subnets</h2>
            <table>
                <tr><th>AD Sites</th><td>$($sites.Count) Sites definiert</td></tr>
                <tr><th>Defined Subnets</th><td>$($subnets.Count) Subnetze</td></tr>
                <tr><th>Trusts</th><td>$(if($trusts){($trusts | ForEach-Object {$_.Name}) -join ", "}else{"Keine"})</td></tr>
            </table>
            <div style="margin-top:10px; font-size:0.8em; color:#888;">
                Standorte: $(($sites | ForEach-Object {$_.Name}) -join ", ")
            </div>
        </div>

        <div class="card">
            <h2>🛡️ Security Indicators</h2>
            <table>
                <tr><th>Domain Admins</th><td class="critical">$adminCount</td></tr>
                <tr><th>Enterprise Admins</th><td class="critical">$enterpriseAdmins</td></tr>
                <tr><th>Inactive Users (90d)</th><td>$inactiveUsers</td></tr>
                <tr><th>Recycle Bin</th><td>$(if ((Get-ADOptionalFeature -Filter 'name -like "Recycle Bin Feature"').EnabledProperties.Count -gt 0) { "Aktiviert" } else { "DEAKTIVIERT" })</td></tr>
            </table>
        </div>
    </div>

    <div class="card" style="margin-top:20px;">
        <h2>🖥️ Domain Controller Details</h2>
        <table>
            <thead>
                <tr>
                    <th>Host Name</th>
                    <th>IP Address</th>
                    <th>Site</th>
                    <th>OS Version</th>
                    <th>GC</th>
                    <th>Type</th>
                </tr>
            </thead>
            <tbody>
                $(foreach ($dc in $dcs) {
                    "<tr>
                        <td><b>$($dc.Name)</b></td>
                        <td>$($dc.IPv4Address)</td>
                        <td>$($dc.Site)</td>
                        <td>$($dc.OperatingSystem)</td>
                        <td>$($dc.IsGlobalCatalog ? "✅" : "❌")</td>
                        <td>$($dc.IsReadOnly ? "RODC" : "RWDC")</td>
                    </tr>"
                } -join "")
            </tbody>
        </table>
    </div>

</body>
</html>
"@

# 3. SPEICHERN & ÖFFNEN
$html | Set-Content $reportPath -Encoding utf8
Start-Process $reportPath
Write-Host "Der optimierte Onepager wurde erstellt und geöffnet." -ForegroundColor Green
