# ObsidianPushOnExit.ps1
# Installed from installer/setup.ps1 (template). UTF-8 without BOM recommended.
#
# After Obsidian exits, runs the same Git sync as ObsidianAutoSync.ps1
# (status / pull --rebase / add -A / commit or amend / push).
# Launches Obsidian with -ArgumentList vault path (single session, no restart loop).
#
# Requires Git 2.9+ for: git pull --rebase --autostash (local edits before pull).
#
# Commit strategy:
#   When $UseAmendCommitStrategy is $true (set at install time):
#     - If the latest commit message starts with obsidian-autosync:, amend and force-push with lease
#     - Otherwise: pull --rebase --autostash, new commit, push
#   When $UseAmendCommitStrategy is $false:
#     - Always: pull --rebase --autostash, new commit, push (no --amend / no force-push)
#
# Usage:
#   .\ObsidianPushOnExit.ps1
#   .\ObsidianPushOnExit.ps1 -VaultPath "D:\Notes"
#   .\ObsidianPushOnExit.ps1 -VaultPath "D:\Notes" -ObsidianExe "C:\...\Obsidian.exe"
#
# Shortcut (no console): use ObsidianPushOnExit-Hidden.vbs in the same folder, or:
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "...\ObsidianPushOnExit.ps1" -VaultPath "..."

param(
    [string]$VaultPath   = __VAULT_PATH__,
    [string]$ObsidianExe = ""   # empty = auto-detect
)

$AUTOSYNC_PREFIX = "obsidian-autosync:"

# Set by installer: $true = allow commit --amend when last commit is autosync; $false = always new commit + push
$UseAmendCommitStrategy = __USE_AMEND_MODE__

$LogDir  = Join-Path $PSScriptRoot "logs"
$LogFile = Join-Path $LogDir ("pushonexit_{0}.log" -f (Get-Date -Format "yyyyMM"))
if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir | Out-Null }

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts][$Level] $Message"
    Add-Content -Path $LogFile -Value $line
    Write-Host $line
}

function Show-Toast {
    param([string]$Text)
    try {
        Add-Type -AssemblyName System.Runtime.WindowsRuntime
        [Windows.UI.Notifications.ToastNotificationManager,Windows.UI.Notifications,ContentType=WindowsRuntime] | Out-Null
        [Windows.UI.Notifications.ToastNotification,Windows.UI.Notifications,ContentType=WindowsRuntime]        | Out-Null
        $template = [Windows.UI.Notifications.ToastTemplateType]::ToastText01
        $xml      = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent($template)
        $xml.GetElementsByTagName("text").Item(0).AppendChild(
            $xml.CreateTextNode($Text)
        ) | Out-Null
        $toast = [Windows.UI.Notifications.ToastNotification]::new($xml)
        [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("ObsidianPushOnExit").Show($toast)
    } catch { }
}

function Find-ObsidianExe {
    # On 32-bit PowerShell, $env:PROGRAMFILES may point to Program Files (x86); also check ProgramW6432 and fixed paths for 64-bit installs.
    $pf86 = ${env:ProgramFiles(x86)}
    $candidates = @(
        "C:\Program Files\Obsidian\Obsidian.exe"
    )
    if ($env:ProgramW6432) {
        $candidates += ($env:ProgramW6432 + "\Obsidian\Obsidian.exe")
    }
    $candidates += @(
        $env:LOCALAPPDATA + "\Obsidian\Obsidian.exe",
        $env:PROGRAMFILES + "\Obsidian\Obsidian.exe",
        $(if ($pf86) { Join-Path $pf86 "Obsidian\Obsidian.exe" } else { $null })
    )
    foreach ($path in $candidates) {
        if ($path -and (Test-Path -LiteralPath $path)) { return $path }
    }
    $lnk = Get-ChildItem "$env:APPDATA\Microsoft\Windows\Start Menu" -Recurse `
               -Filter "Obsidian.lnk" -ErrorAction SilentlyContinue |
           Select-Object -First 1
    if ($lnk) {
        $shell  = New-Object -ComObject WScript.Shell
        $target = $shell.CreateShortcut($lnk.FullName).TargetPath
        if (Test-Path $target) { return $target }
    }
    return $null
}

function Test-LastCommitIsAutoSync {
    $msg = & git log -1 --pretty=%s 2>&1
    return ($LASTEXITCODE -eq 0 -and $msg -like "$AUTOSYNC_PREFIX*")
}

function Invoke-Sync {
    Write-Log "Starting sync..."
    Push-Location -LiteralPath $VaultPath
    try {
        $status = & git status --porcelain 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Log "git status failed: $status" "ERROR"; return
        }
        if (-not $status) {
            Write-Log "No changes; skipping sync."; return
        }
        Write-Log "Changes detected; starting commit..."

        if ($UseAmendCommitStrategy) {
            $isAmend = Test-LastCommitIsAutoSync
        } else {
            $isAmend = $false
        }
        $commitMsg = "$AUTOSYNC_PREFIX {0}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss")

        if (-not $isAmend) {
            # Pull before add/commit; --autostash temporarily stashes local changes (Git 2.9+)
            $pullOut = & git pull --rebase --autostash 2>&1
            if ($LASTEXITCODE -ne 0) {
                Write-Log "git pull --rebase --autostash failed: $pullOut" "WARN"
            }
        }

        $out = & git add -A 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Log "git add failed: $out" "ERROR"; return
        }

        if ($isAmend) {
            $out = & git commit --amend -m $commitMsg 2>&1
            if ($LASTEXITCODE -ne 0) {
                Write-Log "git commit --amend failed: $out" "ERROR"; return
            }
            Write-Log "Commit amended: $commitMsg"
            $useForce = $true
        }
        else {
            $out = & git commit -m $commitMsg 2>&1
            if ($LASTEXITCODE -ne 0) {
                Write-Log "git commit failed: $out" "ERROR"; return
            }
            Write-Log "New commit created: $commitMsg"
            $useForce = $false
        }

        if ($useForce) {
            $out = & git push --force-with-lease 2>&1
        } else {
            $out = & git push 2>&1
        }

        if ($LASTEXITCODE -ne 0) {
            Write-Log "git push failed: $out" "ERROR"
        }
        else {
            $mode = if ($useForce) { "amend + force-push" } else { "push" }
            Write-Log "Sync complete ($mode)"
            Show-Toast "Obsidian vault synced"
        }
    }
    catch {
        Write-Log "Unexpected error: $_" "ERROR"
    }
    finally {
        Pop-Location
    }
}

if (-not (Test-Path -LiteralPath $VaultPath)) {
    Write-Log "Vault not found: $VaultPath" "ERROR"; exit 1
}
if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Write-Log "git not found. Install Git for Windows." "ERROR"; exit 1
}

if (-not $ObsidianExe) { $ObsidianExe = Find-ObsidianExe }
if (-not $ObsidianExe -or -not (Test-Path -LiteralPath $ObsidianExe)) {
    Write-Log "Obsidian.exe not found. Specify -ObsidianExe." "ERROR"
    exit 1
}
Write-Log "Obsidian.exe: $ObsidianExe"

Write-Log "ObsidianPushOnExit started"
Write-Log ("Vault: {0}" -f $VaultPath)
Write-Log ("Amend-when-autosync mode: {0}" -f $UseAmendCommitStrategy)

Write-Log "Starting Obsidian..."
$launchStart = Get-Date
$started = Start-Process -FilePath $ObsidianExe -ArgumentList @($VaultPath) `
    -WorkingDirectory $VaultPath -PassThru

$proc = $null
for ($i = 0; $i -lt 60; $i++) {
    if ($null -ne $started) {
        try {
            $started.Refresh()
            if (-not $started.HasExited) {
                $proc = $started
                break
            }
        } catch {
            $started = $null
        }
    }

    $candidates = Get-Process -Name "Obsidian" -ErrorAction SilentlyContinue |
        Where-Object {
            try { $_.StartTime -ge $launchStart.AddSeconds(-2) } catch { $false }
        }
    if ($candidates) {
        $proc = $candidates | Sort-Object StartTime | Select-Object -Last 1
        break
    }

    Start-Sleep -Seconds 1
}

if (-not $proc) {
    Write-Log "Obsidian process not found; launch may have failed." "WARN"
    exit 1
}

Write-Log "Obsidian started (PID: $($proc.Id)); waiting for exit..."

$proc.WaitForExit()

$exitCode = $proc.ExitCode
if ($null -eq $exitCode) { $exitCode = "n/a" }
Write-Log ("Obsidian exited (ExitCode: {0})" -f $exitCode)

Invoke-Sync
