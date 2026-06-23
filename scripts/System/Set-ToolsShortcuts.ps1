# 1. Pfade definieren
$desktopPath = [System.IO.Path]::Combine($env:USERPROFILE, "Desktop", "AD-tools")
$categories = @("GUI", "CMD", "PS")

# 2. Ordnerstruktur erstellen
if (Test-Path $desktopPath) { Remove-Item $desktopPath -Recurse -Force }
foreach ($cat in $categories) { New-Item -Path "$desktopPath\$cat" -ItemType Directory -Force }

$shell = New-Object -ComObject WScript.Shell

# 3. Tool-Liste (Fokus auf Terminal & Konsolen)
$tools = @(
    @{Cat="CMD"; K="REP"; N="Repadmin"; I="AD Replication Diagnostics"; T="cmd.exe"; A="/k repadmin /showrepl"},
    @{Cat="CMD"; K="DCD"; N="DCDiag"; I="DC Health Check"; T="cmd.exe"; A="/k dcdiag /v"},
    @{Cat="CMD"; K="NLT"; N="NLTest"; I="Test Domain Trusts and Channels"; T="cmd.exe"; A="/k nltest /query"},
    @{Cat="CMD"; K="GPR"; N="GPResult"; I="Analyze RSoP Policy"; T="cmd.exe"; A="/k gpresult /R"},
    @{Cat="CMD"; K="DNS"; N="DNSCmd"; I="DNS Management CLI"; T="cmd.exe"; A="/k dnscmd /?"},
    @{Cat="PS"; K="ADPS"; N="AD Module"; I="Load AD PowerShell Module"; T="powershell.exe"; A="-NoExit -Command Import-Module ActiveDirectory"},
    @{Cat="PS"; K="EXCH"; N="Exchange Shell"; I="Exchange Online and On-Prem Management"; T="powershell.exe"; A="-NoExit -Command Write-Host 'Connect-ExchangeOnline' -FG Cyan"},
    @{Cat="PS"; K="CLEAN"; N="Search-ADAccount"; I="Find disabled and expired accounts"; T="powershell.exe"; A="-NoExit -Command Search-ADAccount -AccountDisabled -UsersOnly"},
    @{Cat="GUI"; K="ADUC"; N="AD Users and Computers"; I="Classic Management"; T="dsa.msc"},
    @{Cat="GUI"; K="ADAC"; N="AD Admin Center"; I="Modern GUI with PS History"; T="dsac.exe"}
)

# 4. Verknüpfungen erstellen (mit Name-Sanitizing)
foreach ($t in $tools) {
    $rawName = "$($t.K) - $($t.N) - $($t.I)"
    # Entfernt Zeichen wie / \ : * ? " < > |
    $safeName = $rawName -replace '[\\\/\:\*\?\"<>\|]', '-' 
    
    $lnkPath = Join-Path -Path "$desktopPath\$($t.Cat)" -ChildPath "$safeName.lnk"
    $shortcut = $shell.CreateShortcut($lnkPath)
    $shortcut.TargetPath = $t.T
    if ($t.A) { $shortcut.Arguments = $t.A }
    $shortcut.Save()
}

# 5. HTML Readme Erstellung
$css = "<style>body{font-family:sans-serif;background:#1e1e1e;color:#eee;padding:20px}table{width:100%;border-collapse:collapse}th,td{padding:10px;border:1px solid #444;text-align:left}th{background:#0078d4;color:white}tr:hover{background:#333}</style>"

foreach ($cat in $categories) {
    $htmlPath = "$desktopPath\$cat\readme.html"
    $rows = ($tools | Where-Object {$_.Cat -eq $cat} | ForEach-Object { "<tr><td><b>$($_.K)</b></td><td>$($_.N)</td><td>$($_.I)</td></tr>" }) -join ""
    $html = "<html><head><meta charset='UTF-8'>$css</head><body><h1>$cat Tools Guide</h1><table><tr><th>Kürzel</th><th>Name</th><th>Info</th></tr>$rows</table></body></html>"
    Set-Content -Path $htmlPath -Value $html -Encoding utf8
}

Write-Host "Ordnerstruktur & HTML-Readmes erfolgreich erstellt!" -ForegroundColor Green
