<# Installs ObsidianPushOnExit.ps1 and ObsidianPushOnExit-Hidden.vbs from template/.
   Asks for the Obsidian vault path (or pass -VaultPath), amend mode (or pass -AmendMode), and writes UTF-8 without BOM.
   Interactive amend prompt: empty input defaults to no (amend disabled).
   Optional Start Menu shortcut (Obsidian-Git.lnk) with Obsidian icon; interactive default no, or pass -CreateStartMenuShortcut.
#>
param(
    [string]$VaultPath = "",
    [string]$OutputDir = "",
    [string]$AmendMode = "",
    [string]$CreateStartMenuShortcut = ""
)

$ErrorActionPreference = "Stop"

#region Paths and template tokens
$InstallerRoot = $PSScriptRoot
$TemplateDir   = Join-Path $InstallerRoot "template"
$PsTemplate    = Join-Path $TemplateDir "ObsidianPushOnExit.ps1"
$VbsTemplate   = Join-Path $TemplateDir "ObsidianPushOnExit-Hidden.vbs"

$Token = @{
    VaultPathPs       = "__VAULT_PATH__"
    UseAmendMode      = "__USE_AMEND_MODE__"
}
#endregion

function Write-SetupLog {
    param([string]$Message, [string]$Level = "INFO")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host ("[{0}] [{1}] {2}" -f $ts, $Level, $Message)
}

function Escape-PowerShellSingleQuotedLiteral {
    param([string]$Path)
    return "'" + $Path.Replace("'", "''") + "'"
}

function Format-VbsQuotedPath {
    param([string]$Path)
    return '"' + $Path.Replace('"', '""') + '"'
}

function Write-Utf8NoBom {
    param([string]$Path, [string]$Content)
    $enc = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($Path, $Content, $enc)
}

function Assert-TemplateFile {
    param([string]$LiteralPath)
    if (-not (Test-Path $LiteralPath)) {
        Write-SetupLog "Template not found: $LiteralPath" "ERROR"
        exit 1
    }
}

function Test-PsTemplatePlaceholders {
    param([string]$Raw)
    if (-not $Raw.Contains($Token.VaultPathPs)) {
        Write-SetupLog "Placeholder $($Token.VaultPathPs) missing in PS template" "ERROR"
        exit 1
    }
    if (-not $Raw.Contains($Token.UseAmendMode)) {
        Write-SetupLog "Placeholder $($Token.UseAmendMode) missing in PS template" "ERROR"
        exit 1
    }
}

function Test-VbsTemplatePlaceholders {
    param([string]$Raw)
    if (-not $Raw.Contains($Token.VaultPathPs)) {
        Write-SetupLog "Placeholder $($Token.VaultPathPs) missing in VBS template" "ERROR"
        exit 1
    }
}

function Find-ObsidianExe {
    $candidates = @(
        (Join-Path $env:LOCALAPPDATA "Programs\Obsidian\Obsidian.exe"),
        (Join-Path $env:LOCALAPPDATA "Obsidian\Obsidian.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "Obsidian\Obsidian.exe"),
        (Join-Path $env:ProgramFiles "Obsidian\Obsidian.exe")
    )
    foreach ($p in $candidates) {
        if ($p -and (Test-Path -LiteralPath $p)) {
            return (Resolve-Path -LiteralPath $p).Path
        }
    }
    $uninstallKeys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    foreach ($pattern in $uninstallKeys) {
        $props = @(Get-ItemProperty $pattern -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*Obsidian*" })
        foreach ($app in $props) {
            if ($app.InstallLocation) {
                $exe = Join-Path $app.InstallLocation.TrimEnd('\') "Obsidian.exe"
                if (Test-Path -LiteralPath $exe) { return (Resolve-Path -LiteralPath $exe).Path }
            }
            if ($app.DisplayIcon) {
                $iconPath = ($app.DisplayIcon -split ",")[0].Trim().Trim('"')
                if ($iconPath -and (Test-Path -LiteralPath $iconPath)) { return (Resolve-Path -LiteralPath $iconPath).Path }
            }
        }
    }
    return $null
}

function Resolve-CreateStartMenuShortcut {
    param([string]$CreateStartMenuShortcut)
    if ($CreateStartMenuShortcut) {
        $t = $CreateStartMenuShortcut.Trim().ToLowerInvariant()
        if ($t -in @("y", "yes", "true", "1")) { return $true }
        if ($t -in @("n", "no", "false", "0")) { return $false }
        Write-SetupLog "Invalid -CreateStartMenuShortcut: $CreateStartMenuShortcut (use yes or no)" "ERROR"
        exit 1
    }
    $ans = Read-Host "Create Start Menu shortcut 'Obsidian-Git' (AppData Programs)? (y/n, default n)"
    if ($null -eq $ans -or [string]::IsNullOrWhiteSpace($ans)) { return $false }
    $t = $ans.Trim().ToLowerInvariant()
    if ($t -in @("y", "yes", "true", "1")) { return $true }
    if ($t -in @("n", "no", "false", "0")) { return $false }
    Write-SetupLog "Unrecognized input; defaulting to no (no shortcut)." "WARN"
    return $false
}

function Install-ObsidianGitStartMenuShortcut {
    param(
        [string]$VbsPath,
        [string]$IconExePath
    )
    $programs = [Environment]::GetFolderPath("Programs")
    $lnkPath = Join-Path $programs "Obsidian-Git.lnk"

    if (Test-Path -LiteralPath $lnkPath) {
        Remove-Item -LiteralPath $lnkPath -Force
        Write-SetupLog "Removed existing shortcut: $lnkPath"
    }

    $wsh = New-Object -ComObject WScript.Shell
    $sc = $wsh.CreateShortcut($lnkPath)
    $sc.TargetPath = $VbsPath
    $sc.WorkingDirectory = Split-Path -Parent $VbsPath
    $sc.Description = "Obsidian vault git push on exit (ObsidianPushOnExit-Hidden.vbs)"
    if ($IconExePath) {
        $sc.IconLocation = "$IconExePath,0"
    }
    $sc.Save()

    Write-SetupLog "Created Start Menu shortcut: $lnkPath"
}

function Resolve-AmendMode {
    param([string]$AmendMode)
    if ($AmendMode) {
        $t = $AmendMode.Trim().ToLowerInvariant()
        if ($t -in @("y", "yes", "true", "1")) { return $true }
        if ($t -in @("n", "no", "false", "0")) { return $false }
        Write-SetupLog "Invalid -AmendMode: $AmendMode (use yes or no)" "ERROR"
        exit 1
    }
    $ans = Read-Host "Use git commit --amend when the last commit is autosync (obsidian-autosync:*)? (y/n, default n)"
    if ($null -eq $ans -or [string]::IsNullOrWhiteSpace($ans)) { return $false }
    $t = $ans.Trim().ToLowerInvariant()
    if ($t -in @("y", "yes", "true", "1")) { return $true }
    if ($t -in @("n", "no", "false", "0")) { return $false }
    Write-SetupLog "Unrecognized input; defaulting to no (amend disabled)." "WARN"
    return $false
}

function Read-VaultPathInput {
    param([string]$VaultPath)
    if ($VaultPath) {
        $vaultInput = $VaultPath.Trim().Trim('"').Trim("'")
        if (-not $vaultInput) {
            Write-SetupLog "-VaultPath is empty." "ERROR"
            exit 1
        }
        if (-not (Test-Path -LiteralPath $vaultInput)) {
            Write-SetupLog "Path does not exist: $vaultInput" "WARN"
        }
        return $vaultInput
    }
    do {
        $vaultInput = Read-Host "Enter the full path to your Obsidian vault"
        $vaultInput = $vaultInput.Trim().Trim('"').Trim("'")
        if (-not $vaultInput) {
            Write-SetupLog "Path is empty. Try again." "WARN"
            continue
        }
        if (-not (Test-Path -LiteralPath $vaultInput)) {
            Write-SetupLog "Path does not exist: $vaultInput" "WARN"
            $cont = Read-Host "Continue anyway? (y/n)"
            if ($cont -ne "y") { continue }
        }
        break
    } while ($true)
    return $vaultInput
}

Assert-TemplateFile -LiteralPath $PsTemplate
Assert-TemplateFile -LiteralPath $VbsTemplate

if (-not $OutputDir) {
    $OutputDir = Join-Path $InstallerRoot "out"
}

Write-SetupLog "Output directory: $OutputDir"

$vaultInput = Read-VaultPathInput -VaultPath $VaultPath
$useAmend   = Resolve-AmendMode -AmendMode $AmendMode
Write-SetupLog ("Amend-when-autosync: {0}" -f $useAmend)

$psQuoted  = Escape-PowerShellSingleQuotedLiteral -Path $vaultInput
$vbsQuoted = Format-VbsQuotedPath -Path $vaultInput

if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
    Write-SetupLog "Created directory: $OutputDir"
}

$psContent = Get-Content -Path $PsTemplate -Raw -Encoding UTF8
Test-PsTemplatePlaceholders -Raw $psContent
$psContent = $psContent.Replace($Token.VaultPathPs, $psQuoted)
$amendLiteral = if ($useAmend) { '$true' } else { '$false' }
$psContent = $psContent.Replace($Token.UseAmendMode, $amendLiteral)

$vbsContent = Get-Content -Path $VbsTemplate -Raw -Encoding UTF8
Test-VbsTemplatePlaceholders -Raw $vbsContent
$vbsContent = $vbsContent.Replace($Token.VaultPathPs, $vbsQuoted)

$psOut  = Join-Path $OutputDir "ObsidianPushOnExit.ps1"
$vbsOut = Join-Path $OutputDir "ObsidianPushOnExit-Hidden.vbs"

Write-Utf8NoBom -Path $psOut -Content $psContent
Write-Utf8NoBom -Path $vbsOut -Content $vbsContent

Write-SetupLog "Wrote: $psOut"
Write-SetupLog "Wrote: $vbsOut"

$doShortcut = Resolve-CreateStartMenuShortcut -CreateStartMenuShortcut $CreateStartMenuShortcut
if ($doShortcut) {
    $obsidianExe = Find-ObsidianExe
    if (-not $obsidianExe) {
        Write-SetupLog "Obsidian.exe not found in common paths; shortcut will use default icon." "WARN"
    } else {
        Write-SetupLog "Using Obsidian icon: $obsidianExe"
    }
    Install-ObsidianGitStartMenuShortcut -VbsPath $vbsOut -IconExePath $obsidianExe
} else {
    Write-SetupLog "Start Menu shortcut skipped. Create a shortcut to ObsidianPushOnExit-Hidden.vbs to run without a console window."
}

Write-SetupLog "Done."
