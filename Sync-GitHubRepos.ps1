# Sync-GitHubRepos.ps1 - Synchronises all GitHub repos to C:\GitHub and manages VS Code workspaces
#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$GitHubUser = 'Miraculix666',
    [string]$BaseDir = 'C:\GitHub',
    [switch]$Silent,
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Helpers ──────────────────────────────────────────────────────────────────
function Write-Step {
    param([int]$N, [int]$Total, [string]$Msg, [string]$Color = 'Cyan')
    if (-not $Silent) { Write-Host "[${N}/${Total}] $Msg" -ForegroundColor $Color }
}
function Write-OK { param([string]$M) if (-not $Silent) { Write-Host "    ✅ $M" -ForegroundColor Green } }
function Write-Warn { param([string]$M) if (-not $Silent) { Write-Host "    ⚠️  $M" -ForegroundColor Yellow } }
function Write-Err { param([string]$M) { Write-Host "    ❌ $M" -ForegroundColor Red } }

# ── Step 1 – Get API Token ────────────────────────────────────────────────────
Write-Step 1 5 "GitHub API Token laden …"

function Get-GitHubToken {
    # 1. Try gh CLI (preferred – it manages token lifecycle itself)
    if (Get-Command gh -ErrorAction SilentlyContinue) {
        try {
            $token = (gh auth token 2>$null)
            if ($token -and $token.Length -gt 10) { return $token.Trim() }
        }
        catch {}
    }
    # 2. Environment variable
    if ($env:GITHUB_TOKEN -and $env:GITHUB_TOKEN.Length -gt 10) {
        return $env:GITHUB_TOKEN
    }
    # 3. Prompt once and store in session env (user already has creds in git credential manager,
    #    but PowerShell can't read them back — that's by design for security)
    if (-not $Silent) {
        Write-Host ""
        Write-Host "  💡 Für die GitHub API wird ein Personal Access Token benötigt." -ForegroundColor Yellow
        Write-Host "  ℹ️  Einmalig eingeben – danach: 'gh auth login --with-token' für Dauerspeicherung." -ForegroundColor DarkGray
        Write-Host ""
        $secure = Read-Host "  GitHub PAT (Scope: repo/public_repo)" -AsSecureString
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
        $plain = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
        if ($plain.Length -gt 10) {
            $env:GITHUB_TOKEN = $plain
            # Also authenticate gh so next run finds it automatically
            $plain | gh auth login --with-token 2>&1 | Out-Null
            return $plain
        }
    }
    return $null
}

$PAT = Get-GitHubToken
if (-not $PAT) {
    Write-Err "Kein GitHub Token verfügbar. Abbruch."
    Write-Host "  Tipp: 'gh auth login' ausführen um gh dauerhaft zu authenticieren." -ForegroundColor Yellow
    exit 1
}
$tokenSource = if ((Get-Command gh -EA SilentlyContinue) -and ((gh auth status 2>&1) -notmatch 'not logged')) { 'gh auth' } else { 'Eingabe' }
Write-OK "API Token geladen (via $tokenSource)"

$headers = @{
    Authorization = "token $PAT"
    Accept        = 'application/vnd.github+json'
}

# ── Step 2 – Fetch all repos from GitHub API ──────────────────────────────────
Write-Step 2 5 "Repos von GitHub API abrufen (User: $GitHubUser) …"

function Get-AllRepos {
    param([string]$User, [hashtable]$Headers)
    $repos = [System.Collections.Generic.List[object]]::new()
    $page = 1
    do {
        $uri = "https://api.github.com/users/$User/repos?per_page=100&page=$page&type=all"
        $response = Invoke-RestMethod -Uri $uri -Headers $Headers -Method Get
        $repos.AddRange($response)
        $page++
    } while ($response.Count -eq 100)
    return $repos
}

try {
    $allRepos = Get-AllRepos -User $GitHubUser -Headers $headers
    Write-OK "$($allRepos.Count) Repos gefunden auf GitHub"
}
catch {
    Write-Err "GitHub API-Fehler: $_"
    exit 1
}

# ── Step 3 – Clone missing / pull existing ────────────────────────────────────
Write-Step 3 5 "Repos synchronisieren …"

$cloned = [System.Collections.Generic.List[string]]::new()
$pulled = [System.Collections.Generic.List[string]]::new()
$skipped = [System.Collections.Generic.List[string]]::new()
$failed = [System.Collections.Generic.List[string]]::new()

# Track clone-URLs already processed (avoid duplicate clones like homelab/homelab-repo)
$processedUrls = @{}

foreach ($repo in $allRepos) {
    $repoName = $repo.name
    $cloneUrl = $repo.clone_url
    $localPath = Join-Path $BaseDir $repoName

    # Skip duplicates (same remote URL already processed)
    if ($processedUrls.ContainsKey($cloneUrl)) {
        Write-Warn "Duplikat übersprungen: $repoName → $cloneUrl (bereits als $($processedUrls[$cloneUrl]))"
        $skipped.Add($repoName)
        continue
    }
    $processedUrls[$cloneUrl] = $repoName

    if (Test-Path $localPath) {
        # Check if it's a git repo
        $existingRemote = git -C $localPath remote get-url origin 2>$null
        if ($existingRemote) {
            if (-not $DryRun) {
                try {
                    git -C $localPath fetch --all --prune -q 2>&1 | Out-Null
                    git -C $localPath pull --ff-only -q 2>&1 | Out-Null
                    $pulled.Add($repoName)
                }
                catch {
                    Write-Warn "Pull fehlgeschlagen für ${repoName}: $_"
                    $failed.Add($repoName)
                }
            }
            else {
                Write-OK "[DRY] Würde pullen: $repoName"
            }
        }
        else {
            Write-Warn "Ordner existiert aber kein Git-Remote: $repoName"
            $skipped.Add($repoName)
        }
    }
    else {
        if (-not $DryRun) {
            try {
                git clone $cloneUrl $localPath -q 2>&1 | Out-Null
                $cloned.Add($repoName)
                Write-OK "Geklont: $repoName"
            }
            catch {
                Write-Err "Klon fehlgeschlagen: ${repoName}: $_"
                $failed.Add($repoName)
            }
        }
        else {
            Write-OK "[DRY] Würde klonen: $repoName → $localPath"
            $cloned.Add($repoName)
        }
    }
}

Write-OK "Sync abgeschlossen — Geklont: $($cloned.Count) | Gepullt: $($pulled.Count) | Übersprungen: $($skipped.Count) | Fehler: $($failed.Count)"

# ── Step 4 – Update all.code-workspace ───────────────────────────────────────
Write-Step 4 5 "all.code-workspace aktualisieren …"

$allWorkspaceFile = Join-Path $BaseDir 'all.code-workspace'

# Load existing workspace for settings/extensions
$existingWorkspace = if (Test-Path $allWorkspaceFile) {
    Get-Content $allWorkspaceFile -Raw | ConvertFrom-Json
}
else {
    [PSCustomObject]@{ folders = @(); settings = @{}; extensions = @{} }
}

# Gather all direct subfolders
$subFolders = Get-ChildItem -Path $BaseDir -Directory |
Where-Object { $_.Name -notmatch '^\.git$' } |
Sort-Object Name

# Build folder entries (relative paths)
$folderEntries = foreach ($dir in $subFolders) {
    [PSCustomObject]@{
        name = $dir.Name
        path = $dir.Name
    }
}

# Rebuild workspace JSON
$newWorkspace = [ordered]@{
    folders    = $folderEntries
    settings   = $existingWorkspace.settings
    extensions = $existingWorkspace.extensions
}

$fCount = $folderEntries.Count
if (-not $DryRun) {
    $newWorkspace | ConvertTo-Json -Depth 10 | Set-Content $allWorkspaceFile -Encoding UTF8
    Write-OK -M "all.code-workspace aktualisiert ($fCount Ordner)"
}
else {
    Write-OK -M "[DRY] all.code-workspace würde $fCount Ordner enthalten"
}

# ── Step 5 – Create per-repo .code-workspace files ───────────────────────────
Write-Step 5 5 "Einzelne .code-workspace Dateien erstellen …"

# Standard settings to embed in every workspace
$wsSettings = $existingWorkspace.settings
if (-not $wsSettings) {
    $wsSettings = [ordered]@{
        "files.encoding"                             = "utf8"
        "files.eol"                                  = "`n"
        "files.trimTrailingWhitespace"               = $true
        "files.insertFinalNewline"                   = $true
        "editor.tabSize"                             = 4
        "editor.detectIndentation"                   = $true
        "editor.rulers"                              = @(120)
        "editor.renderWhitespace"                    = "boundary"
        "editor.formatOnSave"                        = $false
        "git.autofetch"                              = $true
        "git.autofetchPeriod"                        = 180
        "git.confirmSync"                            = $false
        "git.enableSmartCommit"                      = $true
        "terminal.integrated.defaultProfile.windows" = "PowerShell"
    }
}

$wsExtensions = $existingWorkspace.extensions
if (-not $wsExtensions) {
    $wsExtensions = [ordered]@{
        recommendations = @(
            "ms-vscode.powershell",
            "redhat.vscode-yaml",
            "eamodio.gitlens",
            "mhutchie.git-graph",
            "editorconfig.editorconfig",
            "streetsidesoftware.code-spell-checker",
            "streetsidesoftware.code-spell-checker-german"
        )
    }
}

$createdWS = 0
$existingWS = 0

foreach ($dir in $subFolders) {
    $wsFile = Join-Path $BaseDir "$($dir.Name).code-workspace"

    if (Test-Path $wsFile) {
        $existingWS++
        continue  # Don't overwrite existing ones
    }

    $wsContent = [ordered]@{
        folders    = @([ordered]@{ name = $dir.Name; path = $dir.Name })
        settings   = $wsSettings
        extensions = $wsExtensions
    }

    if (-not $DryRun) {
        # Write with relative base = BaseDir (workspace is in BaseDir)
        $wsContent | ConvertTo-Json -Depth 10 | Set-Content $wsFile -Encoding UTF8
        $createdWS++
    }
    else {
        Write-OK "[DRY] Würde erstellen: $($dir.Name).code-workspace"
        $createdWS++
    }
}

Write-OK "Workspace-Dateien: $createdWS neu erstellt, $existingWS bereits vorhanden"

# ── Summary ───────────────────────────────────────────────────────────────────
if (-not $Silent) {
    Write-Host "`n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
    Write-Host " ✅ GitHub Sync abgeschlossen" -ForegroundColor Green
    Write-Host " 📥 Geklont  : $($cloned.Count)  ($($cloned -join ', '))" -ForegroundColor Cyan
    Write-Host " 🔄 Gepullt  : $($pulled.Count)" -ForegroundColor Cyan
    Write-Host " ⏭️  Skippped : $($skipped.Count)" -ForegroundColor DarkGray
    if ($failed.Count -gt 0) {
        Write-Host " ❌ Fehler   : $($failed.Count)  ($($failed -join ', '))" -ForegroundColor Red
    }
    Write-Host " 📁 Workspaces neu: $createdWS" -ForegroundColor Cyan
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n" -ForegroundColor DarkGray
}
