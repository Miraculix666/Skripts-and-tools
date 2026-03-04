param(
  [Parameter(Mandatory=$true)][string]$BootWimPath,
  [Parameter(Mandatory=$true)][string]$DepotWinpePath
)

# Entfernt Anführungszeichen am Anfang/Ende autom. aus Übergabewerten
$BootWimPath = $BootWimPath.Trim('"')
$DepotWinpePath = $DepotWinpePath.Trim('"')

$ErrorActionPreference = 'Stop'

# Arbeitsverzeichnisse
$Work = Join-Path $env:TEMP "winpe_work_$(Get-Date -Format yyyyMMdd_HHmmss)"
$Mount = Join-Path $Work "mount"
$Media = Join-Path $Work "media"
$Sources = Join-Path $Media "sources"

New-Item -ItemType Directory -Force -Path $Mount,$Sources | Out-Null

Write-Host "[*] Kopiere boot.wim in Arbeitsbereich..."
Copy-Item -Path $BootWimPath -Destination (Join-Path $Sources "boot.wim") -Force

# Mount etc. wie gehabt ...

