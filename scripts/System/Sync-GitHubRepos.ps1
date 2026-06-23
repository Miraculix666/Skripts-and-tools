# Sync-GitHubRepos.ps1 - Synchronises all GitHub repos to C:\GitHub and manages VS Code workspaces
#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$GitHubUser = 'Miraculix666',
    [string]$BaseDir = 'C:\GitHub',
    [switch]$Silent,
    [switch]$DryRun,
    [switch]$FullSync
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Load Private Secrets & Run D: Backup ──────────────────────────────────────
$secretsScript = "C:\GitHub\configs\Scripts\Manage-SecretsAndBackup.ps1"
if (Test-Path $secretsScript) {
    & $secretsScript -Load
    & $secretsScript -Backup
}

# ── Helpers ──────────────────────────────────────────────────────────────────
function Write-Step {
    param([int]$N, [int]$Total, [string]$Msg, [string]$Color = 'Cyan')
    if (-not $Silent) { Write-Host "[$N/$Total] $Msg" -ForegroundColor $Color }
}
function Write-OK { param([string]$M) if (-not $Silent) { Write-Host "    [OK] $M" -ForegroundColor Green } }
function Write-Warn { param([string]$M) if (-not $Silent) { Write-Host "    [WARN] $M" -ForegroundColor Yellow } }
function Write-Err { param([string]$M) Write-Host "    [ERR] $M" -ForegroundColor Red }

function Merge-JulesBranchesForRepo {
    param(
        [string]$RepoPath,
        [string]$RepoName
    )
    $ErrorActionPreference = 'Continue'

    # 1. Fetch remote branches
    git -C $RepoPath fetch --all --prune -q 2>$null

    # 2. Get the default branch name (usually main or master)
    $defaultBranch = (git -C $RepoPath symbolic-ref refs/remotes/origin/HEAD 2>$null) -replace '^refs/remotes/origin/'
    if (-not $defaultBranch) {
        $defaultBranch = "main"
        # Try master if main doesn't exist
        $branches = git -C $RepoPath branch --list
        if ($branches -match "master") { $defaultBranch = "master" }
    }

    # 3. Get all remote branches that are not merged into default branch
    $unmerged = @(git -C $RepoPath branch -r --no-merged $defaultBranch 2>$null)
    if (-not $unmerged -or $unmerged.Count -eq 0 -or $unmerged[0].Trim().Length -eq 0) {
        return
    }

    # 4. Filter branches matching Jules PR patterns
    $julesBranchPattern = "^origin/(fix-|test-|perf-|jules-|code-health-|refactor-|security-|performance-)"
    
    $branchesToMerge = @()
    foreach ($line in $unmerged) {
        $b = $line.Trim()
        if ($b -match "^origin/") {
            $branchName = $b -replace '^origin/'
            if ($b -match $julesBranchPattern -and $branchName -ne $defaultBranch) {
                $branchesToMerge += $branchName
            }
        }
    }

    if ($branchesToMerge.Count -eq 0) {
        return
    }

    Write-Host "   [JULES] Found $($branchesToMerge.Count) unmerged Jules branches in $RepoName. Attempting auto-merge..." -ForegroundColor Cyan

    # Save current active branch to restore it later
    $activeBranch = (git -C $RepoPath branch --show-current).Trim()
    if (-not $activeBranch) { $activeBranch = $defaultBranch }

    # Ensure we are on the default branch
    if ($activeBranch -ne $defaultBranch) {
        git -C $RepoPath checkout $defaultBranch -q 2>$null
    }

    foreach ($branch in $branchesToMerge) {
        Write-Host "     > Merging branch: $branch ..." -ForegroundColor Cyan
        
        $commitMsg = "Integrate updates from $branch"
        $mergeOutput = git -C $RepoPath -c core.safecrlf=false merge "origin/$branch" --no-edit -m $commitMsg 2>&1
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "       [OK] Merge successful: $branch" -ForegroundColor Green
            
            # Push the merged default branch to remote
            $pushOutput = git -C $RepoPath push origin $defaultBranch 2>&1
            if ($LASTEXITCODE -eq 0) {
                Write-Host "       [OK] Push successful for $defaultBranch" -ForegroundColor Green
                
                # Delete the remote branch on GitHub
                git -C $RepoPath push origin --delete $branch -q 2>$null
                Write-Host "       [OK] Remote branch deleted: $branch" -ForegroundColor Gray
                
                # Delete local tracking branch if created
                git -C $RepoPath branch -D $branch -q 2>$null
            } else {
                Write-Host "       [FAIL] Push failed for ${defaultBranch}: $pushOutput" -ForegroundColor Red
            }
        } else {
            Write-Host "       [FAIL] Merge conflict in $branch. Aborting merge." -ForegroundColor Yellow
            git -C $RepoPath merge --abort 2>&1 | Out-Null
        }
    }

    # Restore original branch if needed
    if ($activeBranch -ne $defaultBranch) {
        git -C $RepoPath checkout $activeBranch -q 2>$null
    }
}

# ── Step 1 - Get API Token ────────────────────────────────────────────────────
Write-Step 1 6 "GitHub API Token laden ..."

function Get-GitHubToken {
    # Clean up placeholder token from template
    if ($env:GITHUB_TOKEN -eq "your_token_here" -or $env:GITHUB_TOKEN -eq "your_github_token_here") {
        $env:GITHUB_TOKEN = $null
    }

    # 1. Try gh CLI (preferred - it manages token lifecycle itself)
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
    # 3. Prompt once and store in session env
    if (-not $Silent) {
        Write-Host ""
        Write-Host "  * Fuer die GitHub API wird ein Personal Access Token benoetigt." -ForegroundColor Yellow
        Write-Host "  * Einmalig eingeben - danach: 'gh auth login --with-token' fuer Dauerspeicherung." -ForegroundColor DarkGray
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
    Write-Err "Kein GitHub Token verfuegbar. Abbruch."
    Write-Host "  Tipp: 'gh auth login' ausfuehren um gh dauerhaft zu authentifizieren." -ForegroundColor Yellow
    exit 1
}
$tokenSource = if ((Get-Command gh -EA SilentlyContinue) -and (( & { $ErrorActionPreference = 'Continue'; gh auth status 2>&1 } ) -notmatch 'not logged')) { 'gh auth' } else { 'Eingabe' }
Write-OK "API Token geladen (via $tokenSource)"

$headers = @{
    Authorization = "token $PAT"
    Accept        = 'application/vnd.github+json'
}

# ── Step 2 - Fetch all repos from GitHub API ──────────────────────────────────
Write-Step 2 6 "Repos von GitHub API abrufen (User: $GitHubUser) ..."

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

# Map GitHub repos for fast lookup
$githubUrlMap = @{}
$githubNameMap = @{}
foreach ($repo in $allRepos) {
    if ($null -ne $repo.clone_url) { $githubUrlMap[$repo.clone_url.ToLower()] = $repo }
    if ($null -ne $repo.ssh_url) { $githubUrlMap[$repo.ssh_url.ToLower()] = $repo }
    if ($null -ne $repo.html_url) { $githubUrlMap[$repo.html_url.ToLower()] = $repo }
    if ($null -ne $repo.name) { $githubNameMap[$repo.name.ToLower()] = $repo }
}

# ── Step 3 - Auditing & Synchronising Repos ───────────────────────────────────
Write-Step 3 6 "Lokale Ordner analysieren und synchronisieren ..."

# Backup local settings for VS Code and Custom IDE before repo scan/commit
$settingsSyncScript = Join-Path $PSScriptRoot "Sync-VSCodeSettings.ps1"
if (Test-Path $settingsSyncScript) {
    Write-Step 3 6 "Backup fuer VS Code & Custom IDE Einstellungen ausfuehren..."
    try {
        & $settingsSyncScript -Backup
    }
    catch {
        Write-Warn "Fehler beim Sichern der Einstellungen: $_"
    }
}

$subDirs = Get-ChildItem -Path $BaseDir -Directory | Where-Object { $_.Name -notmatch '^\.git$' } | Sort-Object Name

$localRepos = @()
$processedUrls = @{} # Track processed remote URLs to detect duplicates

# First, audit all local directories
foreach ($dir in $subDirs) {
    $localPath = $dir.FullName
    $repoName = $dir.Name
    $isGit = Test-Path (Join-Path $localPath '.git')
    $originUrl = $null
    $gitStatus = "Clean"
    $isDirty = $false
    $ahead = 0
    $behind = 0
    $matchedRepo = $null
    $type = "Non-Git"
    
    if ($isGit) {
        # Get origin URL
        try {
            $originUrl = git -C $localPath remote get-url origin 2>&1
            if ($LASTEXITCODE -eq 0) {
                $originUrl = $originUrl.Trim()
            } else {
                $originUrl = $null
            }
        }
        catch {
            $originUrl = $null
        }

        # Resolve remote mapping
        if ($null -ne $originUrl) {
            $urlKey = $originUrl.ToLower()
            if ($githubUrlMap.ContainsKey($urlKey)) {
                $matchedRepo = $githubUrlMap[$urlKey]
                $type = "Active"
                if ($matchedRepo.name -ne $repoName) {
                    $type = "Moved/Renamed"
                }
            } else {
                $type = "Orphaned"
            }

            # Check if this remote URL was already claimed by another local directory
            if ($githubUrlMap.ContainsKey($urlKey)) {
                if ($processedUrls.ContainsKey($urlKey)) {
                    $type = "Duplicate"
                } else {
                    $processedUrls[$urlKey] = $localPath
                }
            }
        } else {
            $type = "Local Git Only"
        }

        # Check git status (dirty, ahead, behind)
        try {
            $statusOutput = @(git -C $localPath status --porcelain 2>&1)
            if ($LASTEXITCODE -eq 0 -and $statusOutput -and $statusOutput.Count -gt 0 -and $statusOutput[0].Trim().Length -gt 0) {
                $isDirty = $true
                $gitStatus = "Dirty"
            }
        } catch {}

        # Fetch and check ahead/behind
        if (-not $DryRun -and $null -ne $originUrl) {
            try {
                git -C $localPath fetch --prune -q 2>&1 | Out-Null
            } catch {}
        }

        try {
            $aheadCount = git -C $localPath rev-list --count '@{u}..HEAD' 2>&1
            if ($LASTEXITCODE -eq 0) { $ahead = [int]$aheadCount.Trim() }
        } catch {}

        try {
            $behindCount = git -C $localPath rev-list --count 'HEAD..@{u}' 2>&1
            if ($LASTEXITCODE -eq 0) { $behind = [int]$behindCount.Trim() }
        } catch {}

        if ($gitStatus -ne "Dirty") {
            if ($ahead -gt 0 -and $behind -gt 0) { $gitStatus = "Diverged" }
            elseif ($ahead -gt 0) { $gitStatus = "Ahead ($ahead)" }
            elseif ($behind -gt 0) { $gitStatus = "Behind ($behind)" }
            else { $gitStatus = "Clean" }
        } else {
            if ($ahead -gt 0) { $gitStatus += ", Ahead ($ahead)" }
            if ($behind -gt 0) { $gitStatus += ", Behind ($behind)" }
        }
    }

    $localRepos += [PSCustomObject]@{
        Name        = $repoName
        Path        = $localPath
        Type        = $type
        IsGit       = $isGit
        OriginUrl   = $originUrl
        GitStatus   = $gitStatus
        IsDirty     = $isDirty
        Ahead       = $ahead
        Behind      = $behind
        MatchedRepo = $matchedRepo
        ActionTaken = "None"
    }
}

# Perform synchronization actions (Pull active, Clone missing)
$cloned = [System.Collections.Generic.List[string]]::new()
$pulled = [System.Collections.Generic.List[string]]::new()
$skipped = [System.Collections.Generic.List[string]]::new()
$failed = [System.Collections.Generic.List[string]]::new()

# 1. Handle existing local folders
foreach ($repo in $localRepos) {
    if (-not $repo.IsGit) {
        $repo.ActionTaken = "Skipped (Non-Git)"
        $skipped.Add($repo.Name)
        continue
    }

    if ($repo.Type -eq "Orphaned") {
        $repo.ActionTaken = "Skipped (Orphaned/Deleted on GitHub)"
        $skipped.Add($repo.Name)
        continue
    }

    if ($repo.Type -eq "Duplicate") {
        $repo.ActionTaken = "Skipped (Duplicate remote URL)"
        $skipped.Add($repo.Name)
        continue
    }

    if ($FullSync) {
        # ── Full Sync Mode: Commit, Pull/Merge, Push ──────────────────────────
        $oldEAP = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        $actions = [System.Collections.Generic.List[string]]::new()
        $hasErrors = $false

        # 1. Handle uncommitted changes (Dirty)
        if ($repo.IsDirty) {
            if (-not $DryRun) {
                try {
                    # Stage all changes with safecrlf disabled to ignore line-ending warnings
                    git -C $repo.Path -c core.safecrlf=false add -A 2>&1 | Out-Null

                    # Check if there are actual changes staged for commit
                    $staged = @(git -C $repo.Path diff --cached --name-only 2>&1)
                    if ($staged -and $staged.Count -gt 0 -and $staged[0].Trim().Length -gt 0) {
                        # Commit with safecrlf disabled
                        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                        $commitMsg = "Auto-commit: sync on $timestamp"
                        $commitOutput = git -C $repo.Path -c core.safecrlf=false commit -m $commitMsg 2>&1
                        if ($LASTEXITCODE -eq 0) {
                            $actions.Add("Auto-committed")
                            $repo.IsDirty = $false
                            $repo.GitStatus = "Clean"
                        } else {
                            $actions.Add("Commit failed")
                            $hasErrors = $true
                            $repo.ActionTaken = "Commit failed: $commitOutput"
                            $failed.Add($repo.Name)
                        }
                    } else {
                        # No actual staged changes (e.g. only line-ending warning files that are ignored/already matched)
                        $repo.IsDirty = $false
                        $repo.GitStatus = "Clean"
                    }
                }
                catch {
                    $actions.Add("Commit error")
                    $hasErrors = $true
                    $repo.ActionTaken = "Commit error: $_"
                    $failed.Add($repo.Name)
                }
            } else {
                $actions.Add("[DRY] Wuerde Aenderungen committen")
                $repo.IsDirty = $false
                $repo.GitStatus = "Clean"
            }
        }

        # 2. Remote operations (if remote exists and commit succeeded)
        if (-not $hasErrors -and $null -ne $repo.OriginUrl) {
            # Fresh fetch in case remote changed while we were running
            if (-not $DryRun) {
                try {
                    git -C $repo.Path fetch --prune -q 2>&1 | Out-Null
                } catch {}
            }

            # Recalculate ahead & behind
            $ahead = 0
            $behind = 0
            try {
                $aheadCount = git -C $repo.Path rev-list --count '@{u}..HEAD' 2>&1
                if ($LASTEXITCODE -eq 0) { $ahead = [int]$aheadCount.Trim() }
            } catch {}
            try {
                $behindCount = git -C $repo.Path rev-list --count 'HEAD..@{u}' 2>&1
                if ($LASTEXITCODE -eq 0) { $behind = [int]$behindCount.Trim() }
            } catch {}

            # Handle pulling/merging remote changes (Behind)
            if ($behind -gt 0) {
                if (-not $DryRun) {
                    try {
                        # Attempt pull with merge strategy (no rebase, auto merge message)
                        $pullOutput = git -C $repo.Path pull --no-rebase --no-edit 2>&1
                        if ($LASTEXITCODE -eq 0) {
                            $actions.Add("Pulled and Merged")
                            $pulled.Add($repo.Name)
                            # Re-fetch ahead/behind after merge
                            try {
                                $aheadCount = git -C $repo.Path rev-list --count '@{u}..HEAD' 2>&1
                                if ($LASTEXITCODE -eq 0) { $ahead = [int]$aheadCount.Trim() }
                            } catch {}
                            $behind = 0
                        } else {
                            # Pull failed - check if it's a conflict and abort
                            git -C $repo.Path merge --abort 2>&1 | Out-Null
                            $actions.Add("Conflict (Aborted)")
                            $hasErrors = $true
                            $repo.ActionTaken = "Conflict during merge (Aborted)"
                            $failed.Add($repo.Name)
                        }
                    }
                    catch {
                        $actions.Add("Pull failed")
                        $hasErrors = $true
                        $repo.ActionTaken = "Pull failed: $_"
                        $failed.Add($repo.Name)
                    }
                } else {
                    $actions.Add("[DRY] Wuerde pullen and mergen")
                    $pulled.Add($repo.Name)
                    $behind = 0
                }
            }

            # Handle pushing local changes (Ahead)
            if (-not $hasErrors -and $ahead -gt 0) {
                if (-not $DryRun) {
                    try {
                        $pushOutput = git -C $repo.Path push origin 2>&1
                        if ($LASTEXITCODE -eq 0) {
                            $actions.Add("Pushed successfully")
                        } else {
                            $actions.Add("Push failed")
                            $hasErrors = $true
                            $repo.ActionTaken = "Push failed: $pushOutput"
                            $failed.Add($repo.Name)
                        }
                    }
                    catch {
                        $actions.Add("Push error")
                        $hasErrors = $true
                        $repo.ActionTaken = "Push error: $_"
                        $failed.Add($repo.Name)
                    }
                } else {
                    $actions.Add("[DRY] Wuerde pushen")
                }
            }
        }

        # Set final action status
        if (-not $hasErrors) {
            if ($actions.Count -gt 0) {
                $repo.ActionTaken = $actions -join ", "
                if ($repo.IsDirty) {
                    $repo.GitStatus = "Dirty"
                } else {
                    $repo.GitStatus = "Clean"
                }
            } else {
                $repo.ActionTaken = "Up to date"
            }
        }
        $ErrorActionPreference = $oldEAP
        continue
    }

    # ── Standard Mode: Pull only if clean & behind ────────────────────────────
    if ($repo.Type -eq "Local Git Only") {
        $repo.ActionTaken = "Skipped (Local Git Only)"
        $skipped.Add($repo.Name)
        continue
    }

    # For Active or Moved/Renamed repos:
    if ($repo.Behind -gt 0) {
        if ($repo.IsDirty -or $repo.Ahead -gt 0) {
            $repo.ActionTaken = "Skipped Pull (Local changes or ahead)"
            $skipped.Add($repo.Name)
        } else {
            if (-not $DryRun) {
                try {
                    git -C $repo.Path pull --ff-only -q 2>&1 | Out-Null
                    $repo.ActionTaken = "Pulled successfully"
                    $repo.GitStatus = "Clean"
                    $pulled.Add($repo.Name)
                }
                catch {
                    $repo.ActionTaken = "Pull failed: $_"
                    $failed.Add($repo.Name)
                }
            } else {
                $repo.ActionTaken = "[DRY] Wuerde pullen"
                $pulled.Add($repo.Name)
            }
        }
    } else {
        $repo.ActionTaken = "Up to date"
    }

    # Auto-merge jules branches if in FullSync mode
    if ($FullSync -and $repo.IsGit -and $repo.Type -ne "Local Git Only") {
        Merge-JulesBranchesForRepo -RepoPath $repo.Path -RepoName $repo.Name
    }

    # Run daily AD cleanup routine if in scripts-and-tools-pol
    if ($repo.Name -match "scripts-and-tools-pol") {
        $cleanupScript = Join-Path $repo.Path "scripts\system\L-Kennung_IGVP_cleanup.ps1"
        if (Test-Path $cleanupScript) {
            $cleanupLog = Join-Path $repo.Path "scripts\system\last_cleanup_run.txt"
            $today = Get-Date -Format "yyyy-MM-dd"
            $runCleanup = $true
            if (Test-Path $cleanupLog) {
                $lastRun = Get-Content $cleanupLog -Raw
                if ($lastRun.Trim() -eq $today) { $runCleanup = $false }
            }
            if ($runCleanup) {
                Write-Host "   [CLEANUP] Starte taegliche AD-Aufraeumroutine..." -ForegroundColor Cyan
                try {
                    & $cleanupScript -Force 2>&1 | Out-String
                    $today | Set-Content $cleanupLog -Encoding UTF8 -Force
                    Write-Host "   [CLEANUP] Aufraeumroutine beendet." -ForegroundColor Green
                } catch {
                    Write-Warning "Fehler bei AD-Aufraeumroutine: $_"
                }
            }
        }
    }
}

# 2. Clone missing repositories from GitHub
foreach ($ghRepo in $allRepos) {
    $cloneUrl = $ghRepo.clone_url
    $urlKey = $cloneUrl.ToLower()
    
    # Check if we already have a local repo for this clone URL
    $hasLocal = $false
    foreach ($repo in $localRepos) {
        if ($repo.IsGit -and $null -ne $repo.OriginUrl -and $repo.OriginUrl.ToLower() -eq $urlKey) {
            $hasLocal = $true
            break
        }
    }

    if (-not $hasLocal) {
        $localPath = Join-Path $BaseDir $ghRepo.name
        $folderExists = Test-Path $localPath
        $folderNotEmpty = $false
        if ($folderExists) {
            $files = Get-ChildItem -Path $localPath -Force 2>$null
            if ($null -ne $files -and ($files | Measure-Object).Count -gt 0) {
                $folderNotEmpty = $true
            }
        }

        if ($folderNotEmpty) {
            Write-Warn "Klonen uebersprungen fuer $($ghRepo.name): Lokaler Ordner existiert bereits und ist nicht leer!"
            $skipped.Add($ghRepo.name)
            continue
        }

        if (-not $DryRun) {
            try {
                git clone $cloneUrl $localPath -q 2>&1 | Out-Null
                $cloned.Add($ghRepo.name)
                
                # Register new repo in our local tracking
                $localRepos += [PSCustomObject]@{
                    Name        = $ghRepo.name
                    Path        = $localPath
                    Type        = "Active"
                    IsGit       = $true
                    OriginUrl   = $cloneUrl
                    GitStatus   = "Clean"
                    IsDirty     = $false
                    Ahead       = 0
                    Behind      = 0
                    MatchedRepo = $ghRepo
                    ActionTaken = "Cloned successfully"
                }
                Write-OK "Geklont: $($ghRepo.name)"
            }
            catch {
                Write-Err "Klonen fehlgeschlagen: $($ghRepo.name): $_"
                $failed.Add($ghRepo.name)
            }
        }
        else {
            Write-OK "[DRY] Wuerde klonen: $($ghRepo.name) -> $localPath"
            $cloned.Add($ghRepo.name)
        }
    }
}

# Sort local repos list by Name for the output report
$localRepos = $localRepos | Sort-Object Name

# Display audit table
if (-not $Silent) {
    Write-Host "`n Repository Audit & Sync Report:" -ForegroundColor Yellow
    $localRepos | Format-Table -Property Name, Type, GitStatus, ActionTaken -AutoSize
}

Write-OK "Sync abgeschlossen - Geklont: $($cloned.Count) | Gepullt: $($pulled.Count) | Uebersprungen: $($skipped.Count) | Fehler: $($failed.Count)"

# Restore/Apply newly pulled settings for VS Code and Custom IDE
if (Test-Path $settingsSyncScript) {
    Write-Step 3 6 "Wiederherstellen/Uebernehmen der synchronisierten VS Code & Custom IDE Einstellungen..."
    try {
        & $settingsSyncScript -Restore
    }
    catch {
        Write-Warn "Fehler beim Uebernehmen der Einstellungen: $_"
    }
}

# ── Step 4 - Update all.code-workspace ───────────────────────────────────────
Write-Step 4 6 "all.code-workspace aktualisieren ..."

$allWorkspaceFile = Join-Path $BaseDir 'all.code-workspace'

# Load existing workspace for settings/extensions
$existingWorkspace = if (Test-Path $allWorkspaceFile) {
    Get-Content $allWorkspaceFile -Raw | ConvertFrom-Json
}
else {
    [PSCustomObject]@{ folders = @(); settings = [ordered]@{}; extensions = [ordered]@{ recommendations = @() } }
}

# Ensure settings and extensions properties exist to prevent strict mode errors
if ($null -eq $existingWorkspace.PSObject.Properties['settings']) {
    $existingWorkspace | Add-Member -MemberType NoteProperty -Name 'settings' -Value ([ordered]@{})
}
if ($null -eq $existingWorkspace.PSObject.Properties['extensions']) {
    $existingWorkspace | Add-Member -MemberType NoteProperty -Name 'extensions' -Value ([ordered]@{ recommendations = @() })
}

# Build folder entries (only include valid, non-duplicate folders)
# Exclude duplicate folders to keep the workspace clean, or include everything if requested.
# Here we include all unique folders to match "all workspaces in this all Workspaces workspace".
$folderEntries = [System.Collections.Generic.List[object]]::new()
foreach ($repo in $localRepos) {
    # Skip duplicates or directories starting with .
    if ($repo.Type -eq "Duplicate" -or $repo.Name -match '^\.') { continue }
    
    $folderEntries.Add([PSCustomObject]@{
        name = $repo.Name
        path = $repo.Name
    })
}

# Rebuild workspace JSON
$newWorkspace = [ordered]@{
    folders    = $folderEntries.ToArray()
    settings   = $existingWorkspace.settings
    extensions = $existingWorkspace.extensions
}

$fCount = $folderEntries.Count
if (-not $DryRun) {
    $newWorkspace | ConvertTo-Json -Depth 10 | Set-Content $allWorkspaceFile -Encoding UTF8
    Write-OK "all.code-workspace aktualisiert ($fCount Ordner)"
}
else {
    Write-OK "[DRY] all.code-workspace wuerde $fCount Ordner enthalten"
}

# ── Step 5 - Create per-repo .code-workspace files ───────────────────────────
Write-Step 5 6 "Einzelne .code-workspace Dateien erstellen ..."

# Standard settings to embed in every workspace
$wsSettings = $existingWorkspace.settings
$hasSettings = $null -ne $wsSettings
if ($hasSettings) {
    if ($wsSettings -is [System.Collections.IDictionary]) {
        $hasSettings = $wsSettings.Count -gt 0
    } else {
        $hasSettings = ($wsSettings.PSObject.Properties | Measure-Object).Count -gt 0
    }
}
if (-not $hasSettings) {
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
$hasExtensions = $null -ne $wsExtensions
if ($hasExtensions) {
    if ($wsExtensions -is [System.Collections.IDictionary]) {
        $hasExtensions = $wsExtensions.Count -gt 0
    } else {
        $hasExtensions = ($wsExtensions.PSObject.Properties | Measure-Object).Count -gt 0
    }
}
if (-not $hasExtensions) {
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

foreach ($repo in $localRepos) {
    # Skip duplicates
    if ($repo.Type -eq "Duplicate" -or $repo.Name -match '^\.') { continue }
    
    $wsFile = Join-Path $BaseDir "$($repo.Name).code-workspace"

    if (Test-Path $wsFile) {
        $existingWS++
        continue  # Don't overwrite existing ones
    }

    $wsContent = [ordered]@{
        folders    = @([ordered]@{ name = $repo.Name; path = $repo.Name })
        settings   = $wsSettings
        extensions = $wsExtensions
    }

    if (-not $DryRun) {
        $wsContent | ConvertTo-Json -Depth 10 | Set-Content $wsFile -Encoding UTF8
        $createdWS++
    }
    else {
        Write-OK "[DRY] Wuerde erstellen: $($repo.Name).code-workspace"
        $createdWS++
    }
}

Write-OK "Workspace-Dateien: $createdWS neu erstellt, $existingWS bereits vorhanden"

# ── Summary ───────────────────────────────────────────────────────────────────
if (-not $Silent) {
    Write-Host "`n=========================================" -ForegroundColor DarkGray
    Write-Host " [OK] GitHub Sync and Audit abgeschlossen" -ForegroundColor Green
    Write-Host " [IN] Geklont  : $($cloned.Count)  ($($cloned -join ', '))" -ForegroundColor Cyan
    Write-Host " [UP] Gepullt  : $($pulled.Count)  ($($pulled -join ', '))" -ForegroundColor Cyan
    Write-Host " [SK] Uebersprungen : $($skipped.Count)" -ForegroundColor DarkGray
    if ($failed.Count -gt 0) {
        Write-Host " [ERR] Fehler   : $($failed.Count)  ($($failed -join ', '))" -ForegroundColor Red
    }
    Write-Host " [WS] Workspaces neu: $createdWS" -ForegroundColor Cyan
    Write-Host "=========================================`n" -ForegroundColor DarkGray
}

# ── Detailed Failure & Resolution Report ──────────────────────────────────
if ($failed.Count -gt 0) {
    Write-Host "=========================================" -ForegroundColor Yellow
    Write-Host "   [WARN] FEHLERBEHEBUNGS-BERICHT UND EMPFEHLUNGEN" -ForegroundColor Yellow
    Write-Host "=========================================" -ForegroundColor Yellow
    
    foreach ($repo in $localRepos) {
        if ($failed -contains $repo.Name) {
            Write-Host "`n    [DIR] Repository: $($repo.Name)" -ForegroundColor Cyan
            Write-Host "  Pfad: $($repo.Path)" -ForegroundColor DarkGray
            Write-Host "  Fehler: $($repo.ActionTaken)" -ForegroundColor Red
            
            Write-Host "  Empfohlenes Vorgehen:" -ForegroundColor Yellow
            if ($repo.ActionTaken -like "*Conflict*") {
                Write-Host "    Ein Merge-Konflikt ist aufgetreten. Bitte loesen Sie die Konflikte manuell:" -ForegroundColor Gray
                Write-Host "    1. Oeffnen Sie ein Terminal und wechseln Sie in den Ordner:" -ForegroundColor Gray
                Write-Host "       cd `"$($repo.Path)`"" -ForegroundColor Green
                Write-Host "    2. Pruefen Sie den Status und loesen Sie die Konflikte:" -ForegroundColor Gray
                Write-Host "       git status" -ForegroundColor Green
                Write-Host "    3. Nach dem Loesen markieren und committen Sie die Aenderungen:" -ForegroundColor Gray
                Write-Host "       git add ." -ForegroundColor Green
                Write-Host "       git commit -m `"Conflict resolution`"" -ForegroundColor Green
                Write-Host "       git push origin" -ForegroundColor Green
            }
            elseif ($repo.ActionTaken -like "*Push failed*") {
                Write-Host "    Das Pushen der lokalen Aenderungen ist fehlgeschlagen." -ForegroundColor Gray
                Write-Host "    Moegliche Ursache: Neue Remote-Aenderungen oder fehlende Rechte." -ForegroundColor Gray
                Write-Host "    Bitte versuchen Sie:" -ForegroundColor Gray
                Write-Host "       cd `"$($repo.Path)`"" -ForegroundColor Green
                Write-Host "       git pull --rebase" -ForegroundColor Green
                Write-Host "       git push origin" -ForegroundColor Green
            }
            elseif ($repo.ActionTaken -like "*Commit failed*") {
                Write-Host "    Das automatische Committen ist fehlgeschlagen." -ForegroundColor Gray
                Write-Host "    Bitte pruefen Sie den lokalen Git-Status auf Dateisperren oder ungueltige Zeichen:" -ForegroundColor Gray
                Write-Host "       cd `"$($repo.Path)`"" -ForegroundColor Green
                Write-Host "       git status" -ForegroundColor Green
            }
            else {
                Write-Host "    Pruefen Sie den Status manuell im Repository-Verzeichnis:" -ForegroundColor Gray
                Write-Host "       cd `"$($repo.Path)`"" -ForegroundColor Green
                Write-Host "       git status" -ForegroundColor Green
            }
        }
    }
    Write-Host "`n=========================================`n" -ForegroundColor Yellow
}

# ── Generate HTML Report Page and Open in Browser ─────────────────────────────
if ($failed.Count -gt 0 -and -not $DryRun) {
    try {
        $reportPath = "C:\Users\nw0b4746\.gemini\antigravity-ide\sync_issue_report.html"
        
        $htmlContent = @"
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <title>Custom Sync - Fehlerbehebungs-Bericht</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: #0d1117;
            color: #c9d1d9;
            margin: 0;
            padding: 40px;
            display: flex;
            justify-content: center;
        }
        .container {
            max-width: 800px;
            width: 100%;
        }
        .header {
            border-bottom: 1px solid #30363d;
            padding-bottom: 20px;
            margin-bottom: 30px;
        }
        .header h1 {
            color: #ff7b72;
            margin: 0 0 10px 0;
            display: flex;
            align-items: center;
            font-size: 28px;
        }
        .header p {
            color: #8b949e;
            margin: 0;
        }
        .card {
            background-color: #161b22;
            border: 1px solid #30363d;
            border-radius: 8px;
            padding: 24px;
            margin-bottom: 24px;
        }
        .repo-title {
            color: #58a6ff;
            font-size: 20px;
            margin-top: 0;
            margin-bottom: 8px;
        }
        .repo-path {
            font-size: 13px;
            color: #8b949e;
            font-family: monospace;
            margin-bottom: 16px;
        }
        .error-message {
            background-color: rgba(248, 81, 73, 0.1);
            border-left: 4px solid #f85149;
            color: #ff7b72;
            padding: 12px;
            border-radius: 0 4px 4px 0;
            font-family: monospace;
            font-size: 14px;
            margin-bottom: 20px;
        }
        .recommendation-title {
            color: #d29922;
            font-weight: 600;
            font-size: 15px;
            margin-bottom: 10px;
        }
        .code-block {
            background-color: #0d1117;
            border: 1px solid #30363d;
            border-radius: 6px;
            padding: 16px;
            font-family: 'Consolas', 'Courier New', Courier, monospace;
            font-size: 13px;
            color: #e6edf3;
            overflow-x: auto;
            margin-bottom: 12px;
            white-space: pre;
        }
        .code-comment {
            color: #8b949e;
        }
        .copy-btn {
            background-color: #21262d;
            border: 1px solid #30363d;
            color: #c9d1d9;
            padding: 6px 12px;
            border-radius: 6px;
            font-size: 12px;
            cursor: pointer;
            transition: 0.2s;
        }
        .copy-btn:hover {
            background-color: #30363d;
            border-color: #8b949e;
        }
    </style>
    <script>
        function copyCode(id) {
            const code = document.getElementById(id).innerText;
            navigator.clipboard.writeText(code);
            const btn = document.querySelector('[data-id="' + id + '"]');
            const originalText = btn.innerText;
            btn.innerText = "Kopiert!";
            setTimeout(() => { btn.innerText = originalText; }, 2000);
        }
    </script>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>[WARN] Custom Sync Konflikte &amp; Fehler</h1>
            <p>Einige Repositories konnten nicht vollautomatisch synchronisiert werden. Bitte loesen Sie die Probleme manuell.</p>
        </div>
"@
        
        $counter = 1
        foreach ($repo in $localRepos) {
            if ($failed -contains $repo.Name) {
                $codeId = "code-$counter"
                
                $htmlContent += @"
        <div class="card">
            <h2 class="repo-title">[DIR] $($repo.Name)</h2>
            <div class="repo-path">Pfad: $($repo.Path)</div>
            <div class="error-message">Fehler: $($repo.ActionTaken)</div>
            <div class="recommendation-title">Empfohlenes Vorgehen:</div>
"@
                
                if ($repo.ActionTaken -like "*Conflict*") {
                    $htmlContent += @"
            <p>Ein Merge-Konflikt ist aufgetreten. Bitte loesen Sie die Konflikte manuell:</p>
            <div class="code-block" id="$codeId"><span class="code-comment"># 1. In den Ordner wechseln:</span>
cd "$($repo.Path)"
<span class="code-comment"># 2. Konflikte prüfen:</span>
git status
<span class="code-comment"># 3. Konfliktdateien bearbeiten, dann markieren und committen:</span>
git add .
git commit -m "Conflict resolution"
git push origin</div>
            <button class="copy-btn" data-id="$codeId" onclick="copyCode('$codeId')">Befehle kopieren</button>
"@
                }
                elseif ($repo.ActionTaken -like "*Push failed*") {
                    $htmlContent += @"
            <p>Das Pushen ist fehlgeschlagen (evtl. neue Remote-Aenderungen oder fehlende Rechte):</p>
            <div class="code-block" id="$codeId"><span class="code-comment"># 1. In den Ordner wechseln:</span>
cd "$($repo.Path)"
<span class="code-comment"># 2. Aenderungen rebasen und erneut pushen:</span>
git pull --rebase
git push origin</div>
            <button class="copy-btn" data-id="$codeId" onclick="copyCode('$codeId')">Befehle kopieren</button>
"@
                }
                else {
                    $htmlContent += @"
            <p>Pruefen Sie den Status manuell im Repository-Verzeichnis:</p>
            <div class="code-block" id="$codeId">cd "$($repo.Path)"
git status</div>
            <button class="copy-btn" data-id="$codeId" onclick="copyCode('$codeId')">Befehle kopieren</button>
"@
                }
                
                $htmlContent += @"
        </div>
"@
                $counter++
            }
        }
        
        $htmlContent += @"
    </div>
</body>
</html>
"@
        
        $htmlContent | Set-Content $reportPath -Encoding UTF8 -Force
        Start-Process $reportPath
    } catch {
        Write-Warn "HTML-Bericht konnte nicht geoeffnet werden: $_"
    }
}

# ── Step 6 - Antigravity Auto-Update (Once a Day) ─────────────────────────────
$stateDir = "C:\Users\nw0b4746\.gemini\antigravity-ide"
if (Test-Path $stateDir) {
    Write-Step 6 6 "Custom IDE Auto-Update pruefen ..."
    
    $updateLogFile = Join-Path $stateDir "last_update_check.txt"
    $todayDate = Get-Date -Format "yyyy-MM-dd"
    $needsUpdateCheck = $true
    
    if (Test-Path $updateLogFile) {
        $lastCheck = Get-Content $updateLogFile -Raw
        if ($lastCheck.Trim() -eq $todayDate) {
            $needsUpdateCheck = $false
        }
    }
    
    if ($needsUpdateCheck) {
        Write-OK "Fuehre taeglichen Update-Check aus..."
        $todayDate | Set-Content $updateLogFile -Encoding UTF8 -Force
        
        # 1. Update agy CLI
        if (Get-Command agy -ErrorAction SilentlyContinue) {
            try {
                Write-OK "Pruefe agy CLI Updates..."
                if (-not $DryRun) {
                    agy update | Out-Null
                    Write-OK "agy CLI Update-Pruefung abgeschlossen."
                } else {
                    Write-OK "[DRY] Wuerde agy update ausfuehren"
                }
            } catch {
                Write-Warn "Fehler beim agy CLI Update: $_"
            }
        }
        
        # 2. Update Antigravity IDE
        $installerPath = "C:\Users\nw0b4746\AppData\Local\antigravity-updater\installer.exe"
        $appPath = "C:\Users\nw0b4746\AppData\Local\Programs\Antigravity IDE\Antigravity IDE.exe"
        
        if (Test-Path $installerPath) {
            Write-Warn "Neues Update fuer Custom IDE gefunden! Installiere..."
            if (-not $DryRun) {
                try {
                    # Relaunch installer silently and wait for completion
                    Write-OK "Beende Custom IDE und installiere Update im Hintergrund..."
                    Start-Process -FilePath $installerPath -ArgumentList "/S" -Wait
                    
                    # Relaunch IDE after update, restoring all open windows and workspaces automatically
                    if (Test-Path $appPath) {
                        Write-OK "Starte Custom IDE neu..."
                        Start-Process -FilePath $appPath
                    }
                }
                catch {
                    Write-Warn "Fehler beim Installieren des IDE-Updates: $_"
                }
            } else {
                Write-OK "[DRY] Wuerde IDE Update silent ausfuehren (/S) und neu starten"
            }
        } else {
            Write-OK "Custom IDE ist auf dem neuesten Stand (keine pending Updates)."
        }
    } else {
        Write-OK "Update-Pruefung fuer heute bereits abgeschlossen ($lastCheck)."
    }
}

