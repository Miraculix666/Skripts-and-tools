# ============================================================
# Disable Welcome / First Run Screens for all users
# Edge, Office, Teams, Windows + Disable Teams Autostart
# ============================================================

Write-Host "Running with administrative privileges..."

# --- Helper: Create registry key if missing ---
function Ensure-Key {
    param([string]$Path)
    if (!(Test-Path $Path)) {
        New-Item -Path $Path -Force | Out-Null
    }
}

# --- Helper: Set registry value safely ---
function Set-Reg {
    param(
        [string]$Path,
        [string]$Name,
        [Object]$Value,
        [string]$Type = "DWORD"
    )
    Ensure-Key $Path
    Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type -Force
}

# ============================================================
# EDGE
# ============================================================
$edgePolicy = "HKLM:\SOFTWARE\Policies\Microsoft\Edge"
Ensure-Key $edgePolicy
Set-Reg $edgePolicy "HideFirstRunExperience" 1
Set-Reg $edgePolicy "DisableFirstRunExperience" 1

# ============================================================
# OFFICE
# ============================================================
$officeCommon = "HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common"
Ensure-Key $officeCommon
Set-Reg $officeCommon "OptInWizard" 0
Set-Reg $officeCommon "FirstRun" 0
Set-Reg $officeCommon "DisableBootToOfficeStart" 1

# ============================================================
# TEAMS (New Teams)
# ============================================================
$teamsPolicy = "HKLM:\SOFTWARE\Policies\Microsoft\Teams"
Ensure-Key $teamsPolicy
Set-Reg $teamsPolicy "DisableFirstRunExperience" 1

# ============================================================
# WINDOWS OOBE / CONSUMER FEATURES
# ============================================================
$cloudContent = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"
Ensure-Key $cloudContent
Set-Reg $cloudContent "DisableSoftLanding" 1
Set-Reg $cloudContent "DisableConsumerFeatures" 1

$oobe = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OOBE"
Ensure-Key $oobe
Set-Reg $oobe "DisableOOBE" 1

# ============================================================
# DEFAULT USER HIVE
# ============================================================
Write-Host "Loading Default User hive..."

$defaultHivePath = "C:\Users\Default\NTUSER.DAT"
reg load HKU\DefaultUser "$defaultHivePath" | Out-Null

# Edge
Ensure-Key "HKU:\DefaultUser\Software\Microsoft\Edge"
Set-Reg "HKU:\DefaultUser\Software\Microsoft\Edge" "HideFirstRunExperience" 1
Set-Reg "HKU:\DefaultUser\Software\Microsoft\Edge" "DisableFirstRunExperience" 1

# Office
Ensure-Key "HKU:\DefaultUser\Software\Microsoft\Office\16.0\Common"
Set-Reg "HKU:\DefaultUser\Software\Microsoft\Office\16.0\Common" "FirstRun" 0
Set-Reg "HKU:\DefaultUser\Software\Microsoft\Office\16.0\Common" "OptInWizard" 0

# Unload hive
reg unload HKU\DefaultUser | Out-Null
Write-Host "Default User hive updated."

# ============================================================
# DISABLE TEAMS AUTOSTART (ALL USERS)
# ============================================================

# New Teams autostart
$teamsAuto = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
if (Test-Path "$teamsAuto\com.squirrel.Teams.Teams") {
    Remove-ItemProperty -Path $teamsAuto -Name "com.squirrel.Teams.Teams" -Force
}

# Old Teams autostart
$teamsOld = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
if (Test-Path "$teamsOld\Teams") {
    Remove-ItemProperty -Path $teamsOld -Name "Teams" -Force
}

# Disable Teams Machine-Wide Installer autostart
$teamsMWI = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run"
if (Test-Path "$teamsMWI\TeamsMachineInstaller") {
    Remove-ItemProperty -Path $teamsMWI -Name "TeamsMachineInstaller" -Force
}

Write-Host "Teams Autostart disabled."

# ============================================================
Write-Host "All welcome screens disabled successfully."
# ============================================================
