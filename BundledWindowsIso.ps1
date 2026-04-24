<#
.SYNOPSIS
Creates a bundled Windows installation ISO by extracting an input ISO, optionally rebuilding install.wim from selected editions, refreshing the media with Dynamic Update packages (LCU, Setup DU, SafeOS DU, and related prerequisites), and generating a new bootable ISO.

.DESCRIPTION
This script operates on a folder that contains a single Windows ISO (excluding *.bundled.iso). It extracts the ISO into a stable work directory, stages the installation media tree, optionally rebuilds ISO\sources\install.wim from selected WIM indices, refreshes the media using Dynamic Update packages, and builds a new bootable ISO.

Dynamic Update (DU) alignment goal:
- Windows Setup normally contacts Microsoft endpoints early in a feature update or media-based install to acquire Dynamic Update packages, then applies those updates to installation media. These packages can include updates to Setup binaries, SafeOS/WinRE, servicing stack requirements, the latest cumulative update (LCU), and applicable drivers intended for DU.
- In environments where devices should not (or cannot) download these during setup, DU packages can be acquired from Microsoft Update Catalog and applied to the image prior to running Setup.
- This script aims to pre-stage those DU packages into the media so the resulting ISO behaves like a current, self-contained installation source with minimal additional downloads at install/upgrade time.

DU package acquisition:
- If DU-related MSU packages are missing (or if -UpdateMSUs is specified), the script uses MSCatalogLTS to search the Microsoft Update Catalog and download the appropriate packages into a `msus\<category>` directory tree located in the same folder as the source ISO. MSCatalogLTS provides commands for searching and downloading updates from the Microsoft Update Catalog.
- The downloaded packages are saved in `<isoDir>\msus\<category>\` (e.g., msus\SSU\,msus\LCU\, msus\SafeOS\, msus\SetupDU\) so they are reusable across runs and can be applied to the staged media.
- For LCUs: if multiple cumulative updates are found for the detected build, all are downloaded (in oldest-to-newest order) to support checkpoint cumulative update chains; otherwise just the latest is used.
-           if the OnlyLatestLCU option is given, only the latest applicable LCU is downloaded and applied, without checkpoint updates.
- For Setup DU, SafeOS DU, and SSU: the latest applicable package is selected, preferring the same month as the latest LCU.

DU package application targets:
Properly updating installation media involves operating on multiple target images. Microsoft identifies the primary targets as:
- WinPE (boot.wim): used to install/deploy/repair Windows.
- WinRE (winre.wim): recovery environment used for offline repair; based on WinPE.
- Windows OS image(s) (install.wim): one or more Windows editions stored in \sources\install.wim.
- The full media tree: Setup.exe and supporting media files.

This script refreshes the media by applying the DU package types Microsoft documents for Windows installation media:
- Latest Cumulative Update (LCU) (and prerequisites/checkpoints when applicable).
- Setup Dynamic Update (Setup DU): updates setup binaries/files used for feature updates and installs.
- Safe OS Dynamic Update (SafeOS DU): updates the safe operating system used for the recovery environment (WinRE).
- Servicing stack requirements: modern LCUs often embed the servicing stack; separate servicing stack packages may exist only when required.

Checkpoint cumulative updates:
- When the catalog search for the detected build number returns multiple LCU entries, all are downloaded in oldest-to-newest order to ensure the full checkpoint chain is available. The existing KB-ordered application logic applies them in the correct sequence.
- When acquiring DU packages, Microsoft guidance also recommends ensuring DU packages correspond to the same month as the latest cumulative update; if a DU package is not available for that month, use the most recently published version.

Drivers on media:
- The script creates a special folder at the root of the staged ISO named "$WinpeDriver$". Windows Setup can scan this folder for driver INF files during installation.
- Place INF-based drivers (subfolders allowed) under \$WinpeDriver$ in the final media.

SetupConfig + convenience launchers + driver installer:
- The script writes two SetupConfig files into the ISO root:
  - SetupConfig-Upgrade.ini (for in-place upgrades)
  - SetupConfig-Clean.ini (for clean installs)
- The script writes three launcher batch files into the ISO root:
  - Upgrade.cmd: runs setup.exe /auto upgrade and passes SetupConfig-Upgrade.ini via /ConfigFile
  - Clean install.cmd: runs setup.exe /auto clean and passes SetupConfig-Clean.ini via /ConfigFile
  - Install Drivers.cmd: installs drivers from $WinpeDriver$ (if present) using pnputil; intended to be run after the initial installation has completed
- SetupConfig is applied only when setup.exe is launched with /ConfigFile <path>. Microsoft documents that when running setup from media/ISO, you must include /ConfigFile to use SetupConfig.ini.
- /Auto {Clean | Upgrade} controls the automated setup mode.

Index selection:
- If no selection is provided, behavior depends on -UpdateISO:
  - Without -UpdateISO: defaults to ALL indices.
  - With -UpdateISO: defaults to EMPTY selection unless indices are explicitly specified.
- Explicit selection can be made using:
  - -Home, -Pro
  - -Indices with numbers, ranges, labels, wildcard labels (* and ?), or regex labels (re:<pattern>).

UpdateISO behavior:
- -UpdateISO reuses an existing work folder from a prior run.
- If -UpdateISO is specified and NO explicit index selection is provided (-Home/-Pro/-Indices):
  - The script does not rebuild ISO\sources\install.wim
  - The script does not service/refresh the images
  - The script does not apply DU/MSU packages
- To force rebuild (and subsequent servicing/refresh), explicitly specify indices (for example: -Pro or -Indices 6,8,10).

UpdateMSUs behavior:
- -UpdateMSUs forces the DU/MSU download logic via MSCatalogLTS, even if MSU files already exist in the msus subdirectory.
- Without -UpdateMSUs, download occurs only when DU/MSU packages are missing (none present in the msus subdirectory alongside the ISO).

MSU directory layout:
- MSU/CAB packages are downloaded into <isoDir>\msus\<category>\ subdirectories.
- Category subdirectories: SSU, LCU, SafeOS, SetupDU.
- Application targets per category:
  - install.wim (each index): SSU (prerequisites) -> LCU (checkpoint chain)
  - winre.wim (inside install.wim): SafeOS
  - boot.wim (all WinPE indices): SafeOS
  - root of ISO: SetupDU

DryRun behavior:
- With -DryRun, the script completes PREP actions needed to stage the work tree and then prints what would happen for post-PREP actions.

References:
[1] https://learn.microsoft.com/en-us/windows/deployment/update/media-dynamic-update
[2] https://github.com/Marco-online/MSCatalogLTS
[3] https://www.deploymentresearch.com/removing-applications-from-your-windows-11-image-before-and-during-deployment/
[4] https://thedotsource.com/2021/03/16/building-iso-files-with-powershell-7/
[5] https://learn.microsoft.com/en-us/windows-hardware/customize/desktop/unattend/microsoft-windows-pnpcustomizationswinpe-driverpaths
[6] https://community.spiceworks.com/t/autounattend-xml-driver-path-issue-for-windows-11-24h2-and-25h2/1244985
[7] https://github.com/wikijm/PowerShell-AdminScripts/blob/master/Miscellaneous/New-IsoFile.ps1
[8] https://www.winhelponline.com/blog/servicing-stack-diagnosis-dism-sfc/

.PARAMETER Folder
Optional. Folder to process. If omitted, the current directory is used.

.PARAMETER DryRun
Runs PREP actions, then prints what would happen for post-PREP actions.

.PARAMETER CleanWork
Deletes the stable work folder before starting.

.PARAMETER UpdateISO
Reuses an existing work folder. If used without explicit indices, no rebuild/servicing/DU actions occur.

.PARAMETER UpdateMSUs
Forces download of DU/MSU packages into the ISO folder using MSCatalogLTS, even if MSUs already exist.

.PARAMETER OnlyLatestLCU
Forces only the latest applicable LCU is downloaded and applied, without checkpoint updates.

.PARAMETER UseSystemTemp
Creates the work folder under the system temp directory instead of under <Folder>\_WinIsoBundlerWork.

.PARAMETER ShowIndices
Shows available image indices (index and name) and exits.

.PARAMETER Home
Select editions whose normalized label matches "Home" exactly.

.PARAMETER Pro
Select editions whose normalized label matches "Pro" exactly.

.PARAMETER Indices
Comma-separated selector string supporting:
- numbers: 6
- ranges: 3-6, 7-*
- exact labels: "Education N"
- wildcard labels: "*Home*", "* N*"
- regex labels: "re:^Education( N)?$"

.PARAMETER UseADK
Prefer ADK DISM and oscdimg tools when available.

.PARAMETER UseSystem
Force system DISM and PATH oscdimg.

.PARAMETER dism
Explicit path to dism.exe.

.PARAMETER oscdimg
Explicit path to oscdimg.exe.

.PARAMETER ISO
Explicit path to source ISO

.PARAMETER DestISO
Explicit path to destination ISO

.EXAMPLE
PS> .\BundledWindowsIso.ps1 D:\temp -Pro -CleanWork
Builds a bundled ISO using the Pro edition only, deleting any prior work folder first.

.EXAMPLE
PS> .\BundledWindowsIso.ps1 -UpdateISO
Reuses the existing work folder and performs no rebuild/servicing/DU actions (no indices were specified).

.EXAMPLE
PS> .\BundledWindowsIso.ps1 -UpdateISO -Pro
Reuses the existing work folder and forces rebuild/servicing/DU refresh using the Pro edition selection.

.EXAMPLE
PS> .\BundledWindowsIso.ps1 -UpdateMSUs
Forces DU/MSU downloads into the ISO folder using MSCatalogLTS.

.EXAMPLE
PS> .\BundledWindowsIso.ps1 -Indices "* N*"
Selects all N editions (quote required due to space).

.EXAMPLE
PS> .\BundledWindowsIso.ps1 -ShowIndices
Shows install.wim indices and exits.

.NOTES
- Dynamic Update packages can be acquired from Microsoft Update Catalog and applied to installation media prior to running Setup.
- Microsoft documents the DU package categories (LCU, Setup DU, SafeOS DU, servicing stack requirements) and the image targets involved in updating installation media (WinRE, OS image, WinPE, and the media tree).
- Starting with Windows 11, version 24H2, checkpoint cumulative updates might be required as prerequisites for the latest LCU.
- The "$WinpeDriver$" folder at the root of installation media can be used to provide drivers that Setup scans during installation.
#>

# ==============================
$script:Name = "BundledWindowsIso.ps1"
# ==============================

# ==============================
# git information
# ==============================
$GitHash = "7300909"

# ==============================
# Script identity
# ==============================
$script:ScriptPath = $PSCommandPath
if (-not $script:ScriptPath) { $script:ScriptPath = $MyInvocation.MyCommand.Path }

# ==============================
# Configuration
# ==============================
$Config = [ordered]@{
  OutputIsoSuffix        = '.bundled.iso'

  WorkParentSubfolder    = '_WinIsoBundlerWork'
  WorkIsoSubdir          = 'ISO'
  WorkInstallSubdir      = 'INSTALL'
  WorkMountSubdir        = 'MOUNT'
  WorkLogsSubdir         = 'LOGS'
  WorkScratchSubdir      = 'SCRATCH'
  WorkDuSubdir           = 'DU'
  MetaFileName           = 'bundler.meta.txt'

  DriverFolderName       = '$WinpeDriver$'

  RobocopyArgsBase       = @('/E','/R:2','/W:2','/DCOPY:DA','/COPY:DAT')
  RobocopyArgsQuiet      = @('/NFL','/NDL','/NJH','/NJS','/NC','/NS','/NP')

  BootFileBIOS           = 'boot\etfsboot.com'
  BootFileUEFI           = 'efi\microsoft\boot\efisys.bin'
  IsoVolumeLabel         = 'WIN_BUNDLED'
  OscdimgFsArgs          = @('-m','-o','-u2','-udfver102')

  SetupConfigUpgradeName  = 'SetupConfig-Upgrade.ini'
  SetupConfigCleanName    = 'SetupConfig-Clean.ini'
  UpgradeCmdName          = 'Upgrade.cmd'
  CleanCmdName            = 'Clean install.cmd'
  InstallCmdName          = 'Install Drivers.cmd'

  SetupConfigUpgradeLines = @(
    '[SetupConfig]',
    'DynamicUpdate=Disable',
    'ShowOOBE=None',
    'Telemetry=Disable'
  )
  SetupConfigCleanLines   = @(
    '[SetupConfig]',
    'DynamicUpdate=Disable',
    'ShowOOBE=Full',
    'Telemetry=Disable'
  )

  UpgradeCmdTemplate      = @'
@echo off
setlocal
set "SRC=%~dp0"
echo Running in-place upgrade from: %SRC%
"%SRC%setup.exe" /auto upgrade /eula accept /configfile "%SRC%{0}"
echo.
echo If Setup exits immediately, check setup logs under C:\$WINDOWS.~BT\Sources\Panther
endlocal
'@

  CleanCmdTemplate        = @'
@echo off
setlocal
set "SRC=%~dp0"
echo WARNING: This will start a CLEAN install (wipe-and-load) when run from within Windows.
echo Close all apps and ensure you have backups.
echo.
"%SRC%setup.exe" /auto clean /eula accept /configfile "%SRC%{0}"
endlocal
'@

  InstallCmdTemplate        = @'
@echo off
setlocal
set "SRC=%~dp0"
:: Must be run elevated to work
pnputil /add-driver "%SRC%\{0}\*.inf" /subdirs /install
endlocal
'@

  MsusSubdirName            = 'msus'

  CatalogCategoryMatchers = [ordered]@{
    SSU     = '(?i)\bServicing Stack Update\b'
    LCU     = '(?i)\bCumulative Update\b'
    SafeOS  = '(?i)\bSafe OS Dynamic Update\b'
    SetupDU = '(?i)\bSetup Dynamic Update\b'
  }
  CatalogExcludeTitleTokens = @('preview')
  CatalogDownloadAll        = $true

  PreflightCleanupMountPoints = $true
  PreflightUnmountMatchingMountedImages = $true

  # ADK tool discovery
  AdkToolArch            = @('amd64','arm64')
  AdkDeploymentToolsRoot = "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools"
}

# ==============================
# State
# ==============================
$script:State = [ordered]@{
  IsoPath           = $null
  IsoWasMounted     = $false

  WorkBase          = $null
  WorkRoot          = $null
  IsoRoot           = $null
  InstallRoot       = $null
  MountRoot         = $null
  LogsRoot          = $null
  ScratchRoot       = $null
  DuRoot            = $null
  MetaPath          = $null

  DismPath          = $null
  DismLabel         = $null
  OscdimgPath       = $null
  OscdimgLabel      = $null

  OutputIsoPath     = $null
  IsoBaseName       = $null

  StashedInstallWim = $null

  DetectedOS        = $null
  DetectedArch      = $null
  DetectedVersion   = $null
  DetectedBuild     = $null
}

$script:Cancelled   = $false
$script:DryRun      = $false
$script:DryRunPhase = "afterprep" # "prep" or "afterprep"
$script:ChildProcs  = New-Object System.Collections.Generic.List[System.Diagnostics.Process]
$script:CancelHandlerRegistered = $false

# ==============================
# Helpers
# ==============================

function Protect-Token([string]$s) {
  if (-not $s) { return "unknown" }
  $t = $s -replace '[^\w\.-]+','_'
  $t = $t -replace '_+','_'
  return $t.Trim('_')
}

function Show-Usage {
  $name = if ($script:ScriptPath) { Split-Path -Leaf $script:ScriptPath } else { $script:Name }
  Write-Host ""
  Write-Host "$name ($GitHash)" -ForegroundColor Cyan
  Write-Host "Usage:" -ForegroundColor Cyan
  Write-Host "  & '$name' [<Folder>] [-ISO <path>] [-DestISO <path>] [-Home|-Pro|-Indices <spec>] [-CleanWork] [-UpdateISO] [-UpdateMSUs] [-CleanMSUs] [-OnlyLatestLCU] [-DryRun] [-Debug] [-Verbose]" -ForegroundColor Cyan
  Write-Host "  & '$name' -ShowIndices [-ISO <path>]" -ForegroundColor Cyan
  Write-Host ""
  Write-Host "Key Options:" -ForegroundColor Cyan
  Write-Host "  <Folder>          Work directory (default: current directory)" -ForegroundColor Gray
  Write-Host "  -ISO, -SrcISO     Explicit path to input ISO (overrides auto-detect)" -ForegroundColor Gray
  Write-Host "  -DestISO          Explicit path for output bundled ISO" -ForegroundColor Gray
  Write-Host ""
}

function Stop-Script([string]$Message, [int]$Code = 1) {
  Write-Host ""
  Write-Host "ERROR: $Message" -ForegroundColor Red
  Show-Usage
  exit $Code
}

function Assert-Admin {
  $id = [Security.Principal.WindowsIdentity]::GetCurrent()
  $p  = New-Object Security.Principal.WindowsPrincipal($id)
  if (-not $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Stop-Script "Please run PowerShell as Administrator."
  }
}

function Assert-NotCancelled {
  if ($script:Cancelled) { throw [System.OperationCanceledException]::new("Cancelled by user (Ctrl+C).") }
}

function Test-AfterPrepDryRun { return ($script:DryRun -and $script:DryRunPhase -eq 'afterprep') }

function Invoke-Step([string]$What, [scriptblock]$Action) {
  if (Test-AfterPrepDryRun) {
    Write-Host ("[DryRun] {0}" -f $What) -ForegroundColor Yellow
    return $null
  }
  Write-Verbose $What
  return & $Action
}

function New-Folder([string]$Path) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }

function Show-RunBanner {
  Write-Host ""
  Write-Host "Using ISO:" -ForegroundColor Cyan
  Write-Host ("  {0}" -f $script:State.IsoPath) -ForegroundColor Cyan
  Write-Host "Work root:" -ForegroundColor Cyan
  Write-Host ("  {0}" -f $script:State.WorkRoot) -ForegroundColor Cyan
  Write-Host "Detected media:" -ForegroundColor Cyan
  Write-Host ("  OS:       {0}" -f $script:State.DetectedOS) -ForegroundColor Cyan
  if ($script:State.DetectedVersion) { Write-Host ("  Version:  {0}" -f $script:State.DetectedVersion) -ForegroundColor Cyan }
  Write-Host ("  Arch:     {0}" -f $script:State.DetectedArch) -ForegroundColor Cyan
  if ($script:State.DetectedBuild) { Write-Host ("  Build:    {0}" -f $script:State.DetectedBuild) -ForegroundColor Cyan }
  Write-Host ""
}

function Clear-ReadOnlyAttributes {
  param([Parameter(Mandatory=$true)][string]$Path)
  Write-Host ("Clearing Read-only attributes under: {0}" -f $Path) -ForegroundColor Yellow
  Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
    try {
      if ($_.Attributes -band [IO.FileAttributes]::ReadOnly) {
        $_.Attributes = $_.Attributes -band (-bnot [IO.FileAttributes]::ReadOnly)
      }
    } catch {}
  }
}

# ==============================
# Ctrl-C / cancellation handling
# ==============================
function Register-CancelHandler {
  if ($script:CancelHandlerRegistered) { return }

  $handler = [ConsoleCancelEventHandler]{
    param($sender, $e)
    $script:Cancelled = $true
    $e.Cancel = $true
    try { Write-Host "`nCTRL+C detected. Stopping DISM operations and cleaning up..." -ForegroundColor Yellow } catch {}
    try { Clear-Hardened -Aggressive -FromCancel } catch {}
    [Environment]::Exit(1)
  }

  try {
    [Console]::add_CancelKeyPress($handler)
    $script:CancelHandlerRegistered = $true
  } catch {
    try { Write-Host "Failed to install cancel handler" -ForegroundColor Yellow } catch {}
    $script:CancelHandlerRegistered = $false
  }
}

# ==============================
# External process runner
# ==============================
function Invoke-External {
  param(
    [Parameter(Mandatory=$true)][string]$FilePath,
    [Parameter(Mandatory=$true)][string[]]$ArgumentList,
    [string]$StepName = ""
  )

  Assert-NotCancelled

  if (Test-AfterPrepDryRun) {
    Write-Host ("[DryRun] EXEC: {0} {1}" -f $FilePath, ($ArgumentList -join ' ')) -ForegroundColor Yellow
    return 0
  }

  if ($StepName) { Write-Host ("==> {0}" -f $StepName) -ForegroundColor Cyan }

  $p = Start-Process -FilePath $FilePath -ArgumentList $ArgumentList -PassThru -NoNewWindow
  $script:ChildProcs.Add($p) | Out-Null

  try {
    while (-not $p.HasExited) {
      if ($script:Cancelled) {
        try { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue } catch {}
        throw [System.OperationCanceledException]::new("Cancelled by user (Ctrl+C).")
      }
      Start-Sleep -Milliseconds 250
      try { $p.Refresh() } catch {}
    }
  } catch {
    try { $p.Refresh(); if (-not $p.HasExited) { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue } } catch {}
    throw
  }

  try { $p.Refresh() } catch {}
  try { return [int]$p.ExitCode } catch { return 0 }
}

function Stop-TrackedChildren {
  foreach ($p in $script:ChildProcs) {
    try { $p.Refresh(); if (-not $p.HasExited) { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue } } catch {}
  }
}

function Invoke-DismRead {
  param([Parameter(Mandatory=$true)][string[]]$Args)
  Assert-NotCancelled
  $out = & $script:State.DismPath @Args
  return ,$out
}

# ==============================
# DISM mount-state preflight + aggressive cleanup
# ==============================
function Get-MountedWimInfoText {
  try { return ,(& $script:State.DismPath "/Get-MountedWimInfo") } catch { return @() }
}

function ConvertFrom-MountedWimInfo {
  param([string[]]$Lines)

  $entries = @()
  $cur = [ordered]@{}
  foreach ($line in $Lines) {
    if ($line -match '^\s*Mount Dir\s*:\s*(.+)\s*$') { $cur.MountDir = $Matches[1].Trim(); continue }
    if ($line -match '^\s*Image File\s*:\s*(.+)\s*$') { $cur.ImageFile = $Matches[1].Trim(); continue }
    if ($line -match '^\s*Image Index\s*:\s*(\d+)\s*$') { $cur.ImageIndex = [int]$Matches[1]; continue }
    if ($line -match '^\s*Mounted Read/Write\s*:\s*(.+)\s*$') { $cur.ReadWrite = $Matches[1].Trim(); continue }
    if ($line -match '^\s*Status\s*:\s*(.+)\s*$') { $cur.Status = $Matches[1].Trim(); continue }

    if ($line.Trim() -eq '' -and $cur.Contains('MountDir')) {
      $entries += [pscustomobject]$cur
      $cur = [ordered]@{}
    }
  }
  if ($cur.Contains('MountDir')) { $entries += [pscustomobject]$cur }
  return $entries
}

function Dismount-WimMountDir {
  param([Parameter(Mandatory=$true)][string]$MountDir)

  if (-not (Test-Path $MountDir)) { return }
  Write-Host ("Unmounting stale mount (discard): {0}" -f $MountDir) -ForegroundColor Yellow
  $rc = Invoke-External -FilePath $script:State.DismPath -ArgumentList @(
    "/Unmount-Image",
    "/MountDir:$MountDir",
    "/Discard"
  ) -StepName ("DISM Unmount (discard) {0}" -f $MountDir)

  if ($rc -ne 0) {
    Write-Host ("WARNING: Unmount failed (exit {0}) for {1}" -f $rc, $MountDir) -ForegroundColor Yellow
  }
}

function Clear-DismMountPoints {
  Write-Host "Running DISM cleanup for mount points..." -ForegroundColor Yellow
  try {
    $rc = Invoke-External -FilePath $script:State.DismPath -ArgumentList @("/Cleanup-MountPoints") -StepName "DISM Cleanup-MountPoints"
    if ($rc -ne 0) {
      Write-Host ("WARNING: DISM /Cleanup-MountPoints returned {0}" -f $rc) -ForegroundColor Yellow
    }
  } catch {}

  try {
    if (Get-Command Clear-WindowsCorruptMountPoint -ErrorAction SilentlyContinue) {
      Clear-WindowsCorruptMountPoint | Out-Null
    }
  } catch {}
}

function Clear-PreflightDismMounts {
  param(
    [string[]]$RelevantImageFiles = @(),
    [string]$RelevantMountRoot = $null
  )

  if (-not $Config.PreflightCleanupMountPoints) { return }

  Write-Host "Preflight: checking for existing DISM mounted images..." -ForegroundColor Yellow
  $entries = ConvertFrom-MountedWimInfo -Lines (Get-MountedWimInfoText)

  $toUnmount = @()
  foreach ($e in $entries) {
    if (-not $e.MountDir) { continue }

    $match = $false
    if ($RelevantMountRoot -and $e.MountDir.ToLowerInvariant().StartsWith($RelevantMountRoot.ToLowerInvariant())) {
      $match = $true
    }

    if (-not $match -and $RelevantImageFiles -and $e.ImageFile) {
      foreach ($img in $RelevantImageFiles) {
        if ($img -and ($e.ImageFile.Trim('"') -ieq $img.Trim('"'))) { $match = $true; break }
      }
    }

    if ($match) { $toUnmount += $e.MountDir }
  }

  $toUnmount = @($toUnmount | Sort-Object -Unique)
  if ($toUnmount.Count -gt 0 -and $Config.PreflightUnmountMatchingMountedImages) {
    Write-Host ("Preflight: unmounting {0} mounted image(s) related to this run..." -f $toUnmount.Count) -ForegroundColor Yellow
    foreach ($md in $toUnmount) {
      try { Dismount-WimMountDir -MountDir $md } catch {}
    }
  } else {
    Write-Verbose "Preflight: no related mounts to unmount."
  }

  Clear-DismMountPoints
}

function Assert-ImageNotMounted {
  param(
    [Parameter(Mandatory=$true)][string]$ImageFile,
    [Parameter(Mandatory=$true)][int]$Index
  )

  $entries = ConvertFrom-MountedWimInfo -Lines (Get-MountedWimInfoText)
  foreach ($e in $entries) {
    if (-not $e.ImageFile -or -not $e.ImageIndex -or -not $e.MountDir) { continue }
    if ($e.ImageFile.Trim('"') -ieq $ImageFile.Trim('"') -and [int]$e.ImageIndex -eq [int]$Index) {
      Write-Host ("Detected existing mount for {0} (Index {1}) at {2}. Discarding..." -f $ImageFile, $Index, $e.MountDir) -ForegroundColor Yellow
      Dismount-WimMountDir -MountDir $e.MountDir
    }
  }
}

function Stop-DismProcessesForWorkRoot {
  param([string]$WorkRoot)

  if (-not $WorkRoot) { return }
  try {
    $procs = Get-CimInstance Win32_Process -Filter "Name='dism.exe'" -ErrorAction SilentlyContinue
    foreach ($p in $procs) {
      $cmd = $p.CommandLine
      if ($cmd -and ($cmd -like "*$WorkRoot*")) {
        try { Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue } catch {}
      }
    }
  } catch {}
}

# ==============================
# Tool resolution
# ==============================
function Find-AdkToolPath {
  param([Parameter(Mandatory=$true)][ValidateSet('dism','oscdimg')][string]$Tool)
  $root = $Config.AdkDeploymentToolsRoot
  if (-not (Test-Path $root)) { return $null }
  foreach ($arch in $Config.AdkToolArch) {
    $p = if ($Tool -eq 'dism') {
      Join-Path $root (Join-Path $arch 'DISM\dism.exe')
    } else {
      Join-Path $root (Join-Path $arch 'Oscdimg\oscdimg.exe')
    }
    if (Test-Path $p) { return $p }
  }
  return $null
}

function Resolve-Tools {
  param(
    [string]$ExplicitDism,
    [string]$ExplicitOscdimg,
    [switch]$UseADK,
    [switch]$UseSystem
  )

  if ($UseADK -and $UseSystem) { Stop-Script "Use only one of -UseADK or -UseSystem." }

  if ($ExplicitDism) {
    if (-not (Test-Path $ExplicitDism)) { Stop-Script "Specified -dism path not found: $ExplicitDism" }
    $script:State.DismPath = $ExplicitDism; $script:State.DismLabel = "Explicit"
  } else {
    $adk = if (-not $UseSystem) { Find-AdkToolPath dism } else { $null }
    if ($adk) { $script:State.DismPath = $adk; $script:State.DismLabel = "ADK" }
    else { $script:State.DismPath = "$env:windir\System32\dism.exe"; $script:State.DismLabel = "System" }
    if (-not (Test-Path $script:State.DismPath)) { Stop-Script "DISM not found." }
  }

  if ($ExplicitOscdimg) {
    if (-not (Test-Path $ExplicitOscdimg)) { Stop-Script "Specified -oscdimg path not found: $ExplicitOscdimg" }
    $script:State.OscdimgPath = $ExplicitOscdimg; $script:State.OscdimgLabel = "Explicit"
  } else {
    $adk = if (-not $UseSystem) { Find-AdkToolPath oscdimg } else { $null }
    if ($adk) { $script:State.OscdimgPath = $adk; $script:State.OscdimgLabel = "ADK" }
    else {
      $cmd = Get-Command oscdimg.exe -ErrorAction SilentlyContinue
      if (-not $cmd) { Stop-Script "oscdimg.exe not found (install ADK or add to PATH)." }
      $script:State.OscdimgPath = $cmd.Source; $script:State.OscdimgLabel = "PATH"
    }
  }

  Write-Verbose ("Using DISM: {0} ({1})" -f $script:State.DismPath, $script:State.DismLabel)
  Write-Verbose ("Using OSCDIMG: {0} ({1})" -f $script:State.OscdimgPath, $script:State.OscdimgLabel)
}

# ==============================
# Work paths / meta
# ==============================
function Initialize-WorkPaths([string]$FolderPath, [string]$IsoPath, [switch]$UseSystemTemp) {
  $workBase = if ($UseSystemTemp) { [System.IO.Path]::GetTempPath() } else { Join-Path $FolderPath $Config.WorkParentSubfolder }
  New-Item -ItemType Directory -Path $workBase -Force | Out-Null
  $isoBase = [IO.Path]::GetFileNameWithoutExtension($IsoPath)

  $script:State.WorkBase    = $workBase
  $script:State.WorkRoot    = Join-Path $workBase $isoBase
  $script:State.IsoRoot     = Join-Path $script:State.WorkRoot $Config.WorkIsoSubdir
  $script:State.InstallRoot = Join-Path $script:State.WorkRoot $Config.WorkInstallSubdir
  $script:State.MountRoot   = Join-Path $script:State.WorkRoot $Config.WorkMountSubdir
  $script:State.LogsRoot    = Join-Path $script:State.WorkRoot $Config.WorkLogsSubdir
  $script:State.ScratchRoot = Join-Path $script:State.WorkRoot $Config.WorkScratchSubdir
  $script:State.DuRoot      = Join-Path $script:State.WorkRoot $Config.WorkDuSubdir
  $script:State.MetaPath    = Join-Path $script:State.WorkRoot $Config.MetaFileName
  $script:State.IsoBaseName = $isoBase
}

function Write-Meta {
  $lines = @(
    "ISOPath=$($script:State.IsoPath)",
    "IsoBaseName=$($script:State.IsoBaseName)",
    "WorkRoot=$($script:State.WorkRoot)",
    "IsoRoot=$($script:State.IsoRoot)",
    "InstallRoot=$($script:State.InstallRoot)",
    "MountRoot=$($script:State.MountRoot)",
    "LogsRoot=$($script:State.LogsRoot)",
    "ScratchRoot=$($script:State.ScratchRoot)",
    "DuRoot=$($script:State.DuRoot)",
    "StashedInstallWim=$($script:State.StashedInstallWim)",
    "OutputIsoPath=$($script:State.OutputIsoPath)"
  )
  if (Test-AfterPrepDryRun) { Write-Host "[DryRun] Would write meta file." -ForegroundColor Yellow; return }
  Set-Content -Path $script:State.MetaPath -Value $lines -Encoding ASCII
}

function Read-Meta([string]$MetaPath) {
  $h = @{}
  if (-not (Test-Path $MetaPath)) { return $null }
  foreach ($line in Get-Content -LiteralPath $MetaPath) {
    if ($line -match '^\s*([^=]+)=(.*)\s*$') { $h[$Matches[1]] = $Matches[2] }
  }
  return $h
}

# ==============================
# ISO and file ops
# ==============================
function Get-InputIso([string]$FolderPath) {
  $isos = Get-ChildItem -LiteralPath $FolderPath -Filter "*.iso" -File | Where-Object { $_.Name -notlike "*$($Config.OutputIsoSuffix)" }
  if ($isos.Count -ne 1) { Stop-Script "Expected exactly 1 input ISO (excluding *$($Config.OutputIsoSuffix)) in '$FolderPath'. Found $($isos.Count)." }
  return $isos[0].FullName
}

function Mount-Iso([string]$IsoPath) {
  if (Test-AfterPrepDryRun) { return @{ Drive=$null } }
  
  Write-Verbose "Mounting ISO: $IsoPath"
  $img = Mount-DiskImage -ImagePath $IsoPath -PassThru
  $vol = $img | Get-Volume -ErrorAction SilentlyContinue
  
  if (-not $vol -or -not $vol.DriveLetter) {
    Write-Verbose "First mount attempt did not resolve drive letter; retrying..."
    Dismount-DiskImage -ImagePath $IsoPath -ErrorAction SilentlyContinue | Out-Null
    Start-Sleep -Milliseconds 500
    $img = Mount-DiskImage -ImagePath $IsoPath -PassThru
    $vol = $img | Get-Volume -ErrorAction SilentlyContinue
  }
  
  if (-not $vol -or -not $vol.DriveLetter) { 
    Stop-Script "Failed to mount ISO or resolve a drive letter. Check that the ISO is valid and accessible. You may need to manually eject any mounted images in Disk Management." 
  }
  
  $script:State.IsoWasMounted = $true
  Write-Verbose "ISO mounted at drive: $($vol.DriveLetter):"
  return @{ Drive="$($vol.DriveLetter):" }
}

function Copy-IsoContents([string]$SrcDrive, [string]$DstFolder, [string]$IsoPathForDisplay) {
  if (Test-AfterPrepDryRun) { return }

  New-Item -ItemType Directory -Path $DstFolder -Force | Out-Null
  New-Item -ItemType Directory -Path $script:State.LogsRoot -Force | Out-Null
  $logPath = Join-Path $script:State.LogsRoot "robocopy.log"

  $srcDisplay = ($SrcDrive.TrimEnd('\') + '\')
  $dstDisplay = $DstFolder.TrimEnd('\')
  $isoName = [IO.Path]::GetFileName($IsoPathForDisplay)

  Write-Host "Extracting ISO contents (robocopy)..." -ForegroundColor Cyan
  Write-Host ("  ISO:         {0}" -f $isoName) -ForegroundColor Cyan
  Write-Host ("  Mounted as:  {0}" -f $srcDisplay) -ForegroundColor Cyan
  Write-Host ("  Destination: {0}" -f $dstDisplay) -ForegroundColor Cyan
  Write-Host ("  Log:         {0}" -f $logPath) -ForegroundColor Cyan

  $robocopyargs = @($srcDisplay, $dstDisplay, "*.*") + $Config.RobocopyArgsBase
  if ($VerbosePreference -ne 'Continue') { $robocopyargs += $Config.RobocopyArgsQuiet }

  if ($VerbosePreference -ne 'Continue') { robocopy @robocopyargs *> $logPath } else { robocopy @robocopyargs }
  $rc = [int]$LASTEXITCODE
  if ($rc -ge 8) { Stop-Script "Robocopy failed with exit code $rc. See log: $logPath" }
}

function Get-InstallImagePathFromRoot([string]$Root) {
  $sources = Join-Path $Root 'sources'
  $wim = Join-Path $sources 'install.wim'
  $esd = Join-Path $sources 'install.esd'
  if (Test-Path $wim) { return $wim }
  if (Test-Path $esd) { return $esd }
  return $null
}

# ==============================
# WIM helpers
# ==============================
function Get-WimPairs([string]$InstallPath) {
  $out = Invoke-DismRead -Args @("/Get-WimInfo", "/WimFile:$InstallPath")
  if ($LASTEXITCODE -ne 0) { Stop-Script "DISM /Get-WimInfo failed for $InstallPath" }

  $pairs = @()
  $cur = $null
  foreach ($line in $out) {
    if ($line -match '^\s*Index\s*:\s*(\d+)\s*$') { $cur = [int]$Matches[1]; continue }
    if ($cur -ne $null -and $line -match '^\s*Name\s*:\s*(.+)\s*$') {
      $pairs += [pscustomobject]@{ Index=$cur; Name=$Matches[1].Trim() }
      $cur = $null
    }
  }
  if ($pairs.Count -lt 1) { Stop-Script "No indexes found in $InstallPath" }
  return $pairs
}

function Get-IndexNameMap([object[]]$Pairs) {
  $m = @{}
  foreach ($p in $Pairs) { $m[[int]$p.Index] = $p.Name }
  return $m
}

function Show-Indices([string]$InstallPath) {
  $pairs = Get-WimPairs -InstallPath $InstallPath
  foreach ($p in $pairs) { Write-Host ("{0,2} {1}" -f $p.Index, $p.Name) }
}

# ==============================
# Indices selection helpers
# ==============================
function Format-Label([string]$s) {
  if (-not $s) { return "" }
  $t = $s.Trim()
  $t = $t -replace '^(?i)\s*Windows\s+\d+\s+', ''
  $t = $t -replace '\s+', ' '
  return $t.Trim()
}

function Remove-Quotes([string]$s) {
  if ($null -eq $s) { return "" }
  $t = $s.Trim()
  if (($t.StartsWith('"') -and $t.EndsWith('"')) -or ($t.StartsWith("'") -and $t.EndsWith("'"))) { return $t.Substring(1, $t.Length-2) }
  return $t
}

function Test-RegexLabelToken([string]$tok) { (Remove-Quotes $tok).Trim().StartsWith("re:", [System.StringComparison]::OrdinalIgnoreCase) }
function Test-WildcardLabelToken([string]$tok) {
  $t = Remove-Quotes $tok
  if ($t.StartsWith("re:", [System.StringComparison]::OrdinalIgnoreCase)) { return $false }
  return ($t.Contains('*') -or $t.Contains('?'))
}

function Resolve-LabelTokenToIndices {
  param([object[]]$Pairs, [string]$Token)

  $tok = Remove-Quotes $Token
  $items = foreach ($p in $Pairs) { [pscustomobject]@{ Index=[int]$p.Index; Norm=(Format-Label $p.Name) } }

  if (Test-RegexLabelToken $tok) {
    $pat = $tok.Substring(3).Trim()
    if ($pat -eq "") { return @() }
    try {
      $rx = New-Object System.Text.RegularExpressions.Regex ($pat, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
      return @($items | Where-Object { $rx.IsMatch($_.Norm) } | Select-Object -ExpandProperty Index)
    } catch { return @() }
  }

  if (Test-WildcardLabelToken $tok) {
    return @($items | Where-Object { $_.Norm -like $tok } | Select-Object -ExpandProperty Index)
  }

  $key = (Format-Label $tok).ToLowerInvariant()
  return @($items | Where-Object { $_.Norm.ToLowerInvariant() -eq $key } | Select-Object -ExpandProperty Index)
}

function Format-IndicesSpec([string]$Spec) {
  if ($null -eq $Spec) { return "" }
  $s = [string]$Spec
  $s = $s -replace ';', ','
  $s = $s -replace '\s*,\s*', ','
  $s = $s.Trim().Trim(',')
  return $s
}

function Split-SelectorTokens([string]$Spec) {
  $tokens = @()
  $cur = ""
  $inQ = $false
  $qChar = ''
  foreach ($ch in $Spec.ToCharArray()) {
    if (-not $inQ -and ($ch -eq '"' -or $ch -eq "'")) { $inQ=$true; $qChar=$ch; $cur+=$ch; continue }
    if ($inQ -and $ch -eq $qChar) { $inQ=$false; $qChar=''; $cur+=$ch; continue }
    if (-not $inQ -and $ch -eq ',') { $tokens += $cur.Trim(); $cur=""; continue }
    $cur += $ch
  }
  if ($cur.Trim() -ne "") { $tokens += $cur.Trim() }
  return $tokens
}

function ConvertFrom-IndicesSpec {
  param([string]$Spec, [int]$MaxIndex, [object[]]$Pairs)

  $selected = @()
  $badTokens = @()
  $badIdx = @()
  $badLabels = @()

  $Spec = Format-IndicesSpec $Spec
  if (-not $Spec) { return @{ Selected=@(); InvalidTokens=@(); InvalidIndices=@(); InvalidLabels=@() } }

  foreach ($raw in (Split-SelectorTokens $Spec)) {
    $tok = $raw.Trim()
    if (-not $tok) { continue }

    if ($tok -eq '*') { $selected += 1..$MaxIndex; continue }

    if ($tok -match '^\d+$') {
      $n = [int]$tok
      if ($n -lt 1 -or $n -gt $MaxIndex) { $badIdx += $n; continue }
      $selected += $n; continue
    }

    if ($tok -match '^(\d+)-(\d+)$') {
      $a=[int]$Matches[1]; $b=[int]$Matches[2]
      if ($a -gt $b -or $a -lt 1 -or $b -gt $MaxIndex) { $badTokens += $raw; continue }
      $selected += ($a..$b); continue
    }

    if ($tok -match '^(\d+)-\*$') {
      $a=[int]$Matches[1]
      if ($a -lt 1 -or $a -gt $MaxIndex) { $badTokens += $raw; continue }
      $selected += ($a..$MaxIndex); continue
    }

    $idxs = Resolve-LabelTokenToIndices -Pairs $Pairs -Token $tok
    if ($idxs.Count -eq 0) { $badLabels += $tok; continue }
    $selected += $idxs
  }

  $selected = $selected | Sort-Object -Unique
  return @{ Selected=$selected; InvalidTokens=$badTokens; InvalidIndices=$badIdx; InvalidLabels=$badLabels }
}

# ==============================
# Install image stash detection
# ==============================
function Initialize-StashedInstallWim {
  param([string]$InstallRoot)

  $wim = Join-Path $InstallRoot 'install.wim'
  $esd = Join-Path $InstallRoot 'install.esd'
  if (Test-Path $wim) { return $wim }
  if (Test-Path $esd) { return $esd }
  return $null
}

# ==============================
# Rebuild install.wim (per-index progress + summary)
# ==============================
function Build-InstallWimFromSelection {
  param(
    [string]$SourceFile,
    [string]$IsoSourcesDir,
    [int[]]$SelectedSourceIndices,
    [int]$MaxSourceIndex
  )

  $dstWim = Join-Path $IsoSourcesDir 'install.wim'
  $dstEsd = Join-Path $IsoSourcesDir 'install.esd'

  Invoke-Step "Remove old ISO sources install.*" {
    if (Test-Path $dstWim) { Remove-Item -Path $dstWim -Force -ErrorAction SilentlyContinue | Out-Null }
    if (Test-Path $dstEsd) { Remove-Item -Path $dstEsd -Force -ErrorAction SilentlyContinue | Out-Null }
  } | Out-Null

  $isWimSource = $SourceFile -like '*.wim'
  $allSelected = ($SelectedSourceIndices.Count -eq $MaxSourceIndex)
  if ($isWimSource -and $allSelected) {
    Write-Host "All indices selected; copying full install.wim to ISO\sources..." -ForegroundColor Cyan
    Invoke-Step "Copy full install.wim (all indices) to ISO\sources" { Copy-Item -Path $SourceFile -Destination $dstWim -Force } | Out-Null

    Write-Host "ISO\sources\install.wim contains:" -ForegroundColor Cyan
    $pairs = Get-WimPairs -InstallPath $dstWim
    foreach ($p in $pairs) { Write-Host ("  {0}: {1}" -f $p.Index, $p.Name) -ForegroundColor Cyan }
    Write-Host ""
    return $dstWim
  }

  $srcPairs = Get-WimPairs -InstallPath $SourceFile
  $nameMap = Get-IndexNameMap -Pairs $srcPairs

  $sourceType = if ($isWimSource) { 'WIM' } else { 'ESD' }
  Write-Host ("Building ISO\sources\install.wim from selected {0} indices..." -f $sourceType) -ForegroundColor Cyan
  Write-Host ("  Source: {0}" -f $SourceFile) -ForegroundColor Cyan
  Write-Host ("  Dest WIM:   {0}" -f $dstWim) -ForegroundColor Cyan
  Write-Host ""

  $destIndex = 0
  foreach ($srcIndex in $SelectedSourceIndices) {
    Assert-NotCancelled
    $destIndex++
    $srcName = $nameMap[[int]$srcIndex]
    if (-not $srcName) { $srcName = "<unknown>" }

    Write-Host ("Exporting source index {0} -> destination index {1}" -f $srcIndex, $destIndex) -ForegroundColor Cyan
    Write-Host ("  Name: {0}" -f $srcName) -ForegroundColor Cyan

    $rc = Invoke-External -FilePath $script:State.DismPath -ArgumentList @(
      "/Export-Image",
      "/SourceImageFile:$SourceFile",
      "/SourceIndex:$srcIndex",
      "/DestinationImageFile:$dstWim",
      "/Compress:max",
      "/CheckIntegrity"
    ) -StepName ("Export-Image {0} (src idx {1} -> dst idx {2})" -f $srcName, $srcIndex, $destIndex)

    if ($rc -ne 0) { Stop-Script "DISM Export-Image failed for SourceIndex $srcIndex (exit $rc)." }

    Write-Host "  Done." -ForegroundColor Cyan
    Write-Host ""
  }

  Write-Host "Rebuilt ISO\sources\install.wim contains:" -ForegroundColor Cyan
  $rebuiltPairs = Get-WimPairs -InstallPath $dstWim
  foreach ($p in $rebuiltPairs) { Write-Host ("  {0}: {1}" -f $p.Index, $p.Name) -ForegroundColor Cyan }
  Write-Host ""

  return $dstWim
}

# ==============================
# SetupConfig + launchers (ISO root)
# ==============================
function Write-ConfigFiles {
  param([Parameter(Mandatory=$true)][string]$IsoRoot)

  $upgradeIni = Join-Path $IsoRoot $Config.SetupConfigUpgradeName
  $cleanIni   = Join-Path $IsoRoot $Config.SetupConfigCleanName

  $upgradeContent = ($Config.SetupConfigUpgradeLines -join "`r`n") + "`r`n"
  $cleanContent   = ($Config.SetupConfigCleanLines   -join "`r`n") + "`r`n"

  Invoke-Step "Write $($Config.SetupConfigUpgradeName)" { Set-Content -Path $upgradeIni -Value $upgradeContent -Encoding ASCII } | Out-Null
  Invoke-Step "Write $($Config.SetupConfigCleanName)"   { Set-Content -Path $cleanIni   -Value $cleanContent   -Encoding ASCII } | Out-Null
}

function Write-Cmds {
  param([Parameter(Mandatory=$true)][string]$IsoRoot)

  $upgradeCmd = Join-Path $IsoRoot $Config.UpgradeCmdName
  $cleanCmd   = Join-Path $IsoRoot $Config.CleanCmdName
  $installCmd = Join-Path $IsoRoot $Config.InstallCmdName

  $upgradeCmdContent = $Config.UpgradeCmdTemplate -f $Config.SetupConfigUpgradeName
  $cleanCmdContent   = $Config.CleanCmdTemplate   -f $Config.SetupConfigCleanName
  $installCmdContent = $Config.InstallCmdTemplate -f $Config.DriverFolderName

  Invoke-Step "Write $($Config.UpgradeCmdName)" { Set-Content -Path $upgradeCmd -Value $upgradeCmdContent -Encoding ASCII } | Out-Null
  Invoke-Step "Write $($Config.CleanCmdName)"   { Set-Content -Path $cleanCmd   -Value $cleanCmdContent   -Encoding ASCII } | Out-Null
  Invoke-Step "Write $($Config.InstallCmdName)" { Set-Content -Path $installCmd -Value $installCmdContent -Encoding ASCII } | Out-Null
}

# ==============================
# Detection (OS, arch, version)
# ==============================
function Get-MediaInfoFromInstallWim {
  param([Parameter(Mandatory=$true)][string]$InstallWim)

  $pairs = Get-WimPairs -InstallPath $InstallWim
  $sampleName = ($pairs | Select-Object -First 1).Name

  $arch = $null
  $build = $null
  $version = $null
  $servicepack = $null

  $out = Invoke-DismRead -Args @("/Get-WimInfo", "/WimFile:$InstallWim", "/Index:1")
  $count = 0
  foreach ($line in $out) {
    $count++
    if ($count -lt 7) { continue } # First 7 lines are header
    if (-not $arch -and $line -match '^\s*Architecture\s*:\s*(.+)\s*$') {
      $arch = $Matches[1].Trim().ToLowerInvariant()
      if ($arch -eq 'amd64') { $arch = 'x64' }
    }
    if (-not $build -and $line -match '^\s*Version\s*:\s*\d+\.\d+\.(\d+)\s*$') {
      $build = $Matches[1]
    }
    if (-not $servicepack -and $line -match '^\s*ServicePack Build\s*:\s*(\d+)\s*$') {
      $servicepack = $Matches[1]
    }
  }

  if ($sampleName -match '(?i)Windows\s+10') {
    $os = "Windows 10"
    if (-not $version -and $build) {
      if ($build -ge 19045) { $version = "22H2" }
      elseif ($build -eq 19044) { $version = "21H2" }
      elseif ($build -eq 19043) { $version = "21H1" }
      elseif ($build -eq 19042) { $version = "20H2" }
      elseif ($build -eq 19041) { $version = "2004" }
      elseif ($build -ge 18363) { $version = "1909" }
      elseif ($build -eq 18362) { $version = "1903" }
      elseif ($build -ge 17763) { $version = "1809" }
      elseif ($build -ge 17134) { $version = "1803" }
      elseif ($build -ge 16299) { $version = "1709" }
      elseif ($build -ge 15063) { $version = "1703" }
      elseif ($build -ge 14393) { $version = "1607" }
      elseif ($build -ge 10586) { $version = "1511" }
      else { $version = "1507" }
    }
  } elseif ($sampleName -match '(?i)Windows\s+11') {
    $os = "Windows 11"
    if (-not $version -and $build) {
      if ($build -ge 28000) { $version = "26H1" }
      elseif ($build -ge 26200) { $version = "25H2" }
      elseif ($build -ge 26100) { $version = "24H2" }
      elseif ($build -ge 22631) { $version = "23H2" }
      elseif ($build -ge 22621) { $version = "22H2" }
      else { $version = "21H1" }
    }
  } elseif ($sampleName -match '(?i)Windows\s+Server') {
    $os = "Windows Server"
  } else {
    Stop-Script "Windows 10, 11, or Server not found. Cannot proceed."
  }

  if ($servicepack) {
    $build = $build + '.' + $servicepack
  }

  if (-not $arch) {
    $b = $script:State.IsoBaseName
    if ($b -match '(?i)(arm64)') { $arch = 'arm64' }
    elseif ($b -match '(?i)(x64|amd64)') { $arch = 'x64' }
    elseif ($b -match '(?i)(x86)') { $arch = 'x86' }
    else { $arch = 'x64' }
  }

  $script:State.DetectedOS = $os
  $script:State.DetectedVersion = $version
  $script:State.DetectedArch = $arch
  $script:State.DetectedBuild = $build
}

# ==============================
# MSCatalogLTS download helpers
# ==============================
function Initialize-MSCatalogLTS {
  $m = Get-Module -ListAvailable -Name MSCatalogLTS -ErrorAction SilentlyContinue
  if (-not $m) {
    Write-Host "MSCatalogLTS module not found; installing from PowerShell Gallery..." -ForegroundColor Cyan
    try {
      try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
      Install-Module -Name MSCatalogLTS -Scope CurrentUser -Force -ErrorAction Stop
    } catch {
      Stop-Script "Failed to install MSCatalogLTS. Install it manually: Install-Module MSCatalogLTS -Scope CurrentUser. Error: $($_.Exception.Message)"
    }
  }
  try {
    Import-Module MSCatalogLTS -Force -ErrorAction Stop
  } catch {
    Stop-Script "Failed to import MSCatalogLTS module. Error: $($_.Exception.Message)"
  }
}

function Test-TitleExcluded([string]$Title) {
  if (-not $Title) { return $true }
  $t = $Title.ToLowerInvariant()
  foreach ($tok in $Config.CatalogExcludeTitleTokens) {
    if ($t -like "*$tok*") { return $true }
  }
  return $false
}

function Get-KBFromTitle([string]$Title) {
  if (-not $Title) { return $null }
  $m = [regex]::Match($Title, '(?i)\bKB(\d{6,8})\b')
  if ($m.Success) { return ("KB{0}" -f $m.Groups[1].Value) }
  return $null
}

function Get-YYYYMMFromTitle([string]$Title) {
  if (-not $Title) { return $null }
  $m = [regex]::Match($Title, '^(?i)(\d{4})-(\d{2})\b')
  if ($m.Success) { return ("{0}-{1}" -f $m.Groups[1].Value, $m.Groups[2].Value) }
  return $null
}

function Get-LocalMsuFiles([string]$FolderPath) {
  return @(Get-ChildItem -LiteralPath $FolderPath -Filter "*.msu" -File -ErrorAction SilentlyContinue | Sort-Object Name)
}

function Clear-MsuFolder([string]$FolderPath) {
  $msuRoot = Join-Path $FolderPath $Config.MsusSubdirName
  if (-not (Test-Path $msuRoot)) { return }
  $existing = @(Get-ChildItem -LiteralPath $msuRoot -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in @('.msu','.cab') })
  if ($existing.Count -lt 1) { return }
  Write-Host ("[CleanMSUs] Removing msus directory ({0} file(s)) in {1}" -f $existing.Count, $msuRoot) -ForegroundColor Yellow
  try { Remove-Item -Path $msuRoot -Recurse -Force -ErrorAction SilentlyContinue } catch {}
}

function Save-CatalogUpdateAllFiles {
  param(
    [Parameter(Mandatory=$true)][object]$CatalogItem,
    [Parameter(Mandatory=$true)][string]$DestinationFolder
  )

  if ($Config.CatalogDownloadAll) {
    $CatalogItem | Save-MSCatalogUpdate -Destination $DestinationFolder -DownloadAll
  } else {
    $CatalogItem | Save-MSCatalogUpdate -Destination $DestinationFolder
  }
}

function Get-CatalogQueryResults {
  param(
    [string]$OsName,
    [string]$OsVersion,
    [string]$Arch,
    [string]$ExtraOptions
  )

  Assert-NotCancelled
  $query =
    "-Descending -IncludeDynamic -AllPages -Search " +
    '"' + "$OsName $OsVersion $Arch" + '"' +
    $(if ($ExtraOptions -and $ExtraOptions.Trim()) { " $ExtraOptions" }) +
    $(if ($DebugSwitch)   { " -Debug" }) +
    $(if ($VerboseSwitch) { " -Verbose" })

  Write-Verbose ("Catalog query attempt: $query")
  try {
    $res = Invoke-Expression "Get-MSCatalogUpdate $query"
  } catch {
    $res = @()
  }
  Write-Verbose ("Catalog query result count for '{0}': {1}" -f $query, $res.Count)
  return $res
}

function Get-CatalogCandidates {
  param(
    [Parameter(Mandatory=$true)][string]$OsName,
    [Parameter(Mandatory=$true)][string]$OsVersion,
    [Parameter(Mandatory=$true)][string]$Arch
  )

  Write-Verbose ("Getting catalog query results")
  # Have to make separate queries for each UpdateType to make sure each is obtained, but any query could return empty
  $res  = Get-CatalogQueryResults $OsName $OsVersion $Arch ('-UpdateType "Cumulative Updates"')
  $res += Get-CatalogQueryResults $OsName $OsVersion $Arch ('-UpdateType "Critical Updates"')
  $res += Get-CatalogQueryResults $OsName $OsVersion $Arch ('-UpdateType "Security Updates"')
  Write-Host ("Catalog query result count: {0}" -f $res.Count)
  return $res
}

function Select-LatestByCategoryPreferMonth {
  param(
    [Parameter(Mandatory=$true)][object[]]$Results,
    [Parameter(Mandatory=$true)][ValidateSet('SSU','LCU','SafeOS','SetupDU')][string]$Category,
    [string]$PreferYYYYMM
  )

  $rx = $Config.CatalogCategoryMatchers[$Category]

  $filtered = @(
    $Results |
      Where-Object { $_.Title } |
      Where-Object { -not (Test-TitleExcluded $_.Title) } |
      Where-Object { $_.Title -match $rx }
  )

  if ($filtered.Count -lt 1) { return $null }

  if ($PreferYYYYMM) {
    $pat = '^' + [regex]::Escape($PreferYYYYMM) + '\b'
    $sameMonth = @($filtered | Where-Object { $_.Title -match $pat })
    if ($sameMonth.Count -gt 0) {
      return ($sameMonth | Sort-Object LastUpdated -Descending | Select-Object -First 1)
    }
  }

  return ($filtered | Sort-Object LastUpdated -Descending | Select-Object -First 1)
}

function Get-AllByCategoryFiltered {
  param(
    [Parameter(Mandatory=$true)][object[]]$Results,
    [Parameter(Mandatory=$true)][ValidateSet('SSU','LCU','SafeOS','SetupDU')][string]$Category
  )

  $rx = $Config.CatalogCategoryMatchers[$Category]

  return @(
    $Results |
      Where-Object { $_.Title } |
      Where-Object { -not (Test-TitleExcluded $_.Title) } |
      Where-Object { $_.Title -match $rx }
  )
}

function Initialize-AllMSUsPresent {
  param(
    [Parameter(Mandatory=$true)][string]$IsoFolder,
    [Parameter(Mandatory=$true)][string]$OsName,
    [Parameter(Mandatory=$true)][string]$OsVersion,
    [Parameter(Mandatory=$true)][string]$OsBuild,
    [Parameter(Mandatory=$true)][string]$Arch,
    [Parameter(Mandatory=$true)][bool]$ForceDownload
  )

  Initialize-MSCatalogLTS

  $msuRoot = Join-Path $IsoFolder $Config.MsusSubdirName
  $existing = @()
  if (Test-Path $msuRoot) {
    $existing = @(Get-ChildItem -LiteralPath $msuRoot -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in @('.msu','.cab') })
  }
  $needDownload = $ForceDownload -or ($existing.Count -lt 1)

  if (-not $needDownload) {
    Write-Verbose ("Updates already present in {0} ({1} file(s)); skipping download. Use -UpdateMSUs or -CleanMSUs to refresh." -f $msuRoot, $existing.Count)
    return [pscustomobject]@{ Selected=[ordered]@{}; Downloaded=@() }
  }

  $results = Get-CatalogCandidates -OsName $OsName -OsVersion $OsVersion -Arch $Arch
  if ($results.Count -lt 1) {
    Write-Host "Catalog search returned no results for $OsName $OsVersion $Arch."
     return [pscustomobject]@{}
  }

  Write-Host ("Search completed: found {0} updates" -f $results.Count) -ForegroundColor Cyan
  Write-Host ""
  if ($VerbosePreference -eq 'Continue') {
    foreach ($r in $results) { Write-Host ("  {0}" -f $r.Title) -ForegroundColor Cyan }
    Write-Host ""
  }

  $selected = [ordered]@{}

  # Collect all LCU candidates for this build. When multiple are present they form a
  # checkpoint cumulative update chain (newer LCU requires an earlier LCU as a prerequisite).
  # Download all of them sorted oldest-first so the full chain is available; just the latest
  # is used when only one is found (no checkpoint chain required).
  $lcuAll = @(Get-AllByCategoryFiltered -Results $results -Category 'LCU')
  if ($lcuAll.Count -lt 1) {
     Write-Host "Could not find any catalog entries for LCU ($OsName $OsVersion $Arch)."
     return [pscustomobject]@{}
  }

  $lcuLatest = ($lcuAll | Select-Object -First 1)
  $lcuYM = (Get-YYYYMMFromTitle $lcuLatest.Title)
  Write-Verbose ("Latest LCU: {0}" -f $lcuLatest.Title)

  if ($lcuAll.Count -gt 1 -and -not [bool]$OnlyLatestLCU) {
    # Multiple LCU entries detected: download the full checkpoint chain oldest-to-newest
    Write-Host ("Found {0} LCU entries for build {1}; downloading full checkpoint chain." -f $lcuAll.Count, $OsBuild) -ForegroundColor Cyan
    $selected['LCU'] = @($lcuAll | ForEach-Object {
      [pscustomobject]@{ Category='LCU'; Title=$_.Title; KB=(Get-KBFromTitle $_.Title); YM=(Get-YYYYMMFromTitle $_.Title); Item=$_ }
    })
  } else {
    # Single LCU entry or OnlyLatestLCU flag is set: no checkpoint chain required
    $selected['LCU'] = @($lcuLatest | ForEach-Object {
      [pscustomobject]@{ Category='LCU'; Title=$_.Title; KB=(Get-KBFromTitle $_.Title); YM=(Get-YYYYMMFromTitle $_.Title); Item=$_ }
    })
  }

  # Non-LCU categories always use the single latest applicable package.
  # LCU may include multiple entries (checkpoint chain); others are always a single entry.
  foreach ($cat in @('SSU','SafeOS','SetupDU')) {
    $pick = Select-LatestByCategoryPreferMonth -Results $results -Category $cat -PreferYYYYMM $lcuYM
    if ($pick) {
      $selected[$cat] = @([pscustomobject]@{
        Category=$cat; Title=$pick.Title; KB=(Get-KBFromTitle $pick.Title); YM=(Get-YYYYMMFromTitle $pick.Title); Item=$pick
      })
    }
  }

  $before = @()
  if (Test-Path $msuRoot) {
    $before = @(Get-ChildItem -LiteralPath $msuRoot -Recurse -File -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName)
  }

  if ($VerbosePreference -eq 'Continue') {
    foreach ($c in $selected.Keys) {
      foreach ($entry in $selected[$c]) { Write-Host ("  {0}: {1}" -f $c, $entry.Title) -ForegroundColor Cyan }
    }
    Write-Host ""
  }

  Write-Host ("Downloading updates to: {0}" -f $msuRoot) -ForegroundColor Cyan
  foreach ($cat in $selected.Keys) {
    $catDir = Join-Path $msuRoot $cat
    if (-not (Test-Path $catDir)) {
      New-Item -ItemType Directory -Path $catDir -Force | Out-Null
    }
    foreach ($entry in $selected[$cat]) {
      Assert-NotCancelled
      $t = $entry.Title
      Write-Host ("Downloading ({0}): {1}" -f $cat, $t) -ForegroundColor Cyan
      try {
        Save-CatalogUpdateAllFiles -CatalogItem $entry.Item -DestinationFolder $catDir
      } catch {
        Stop-Script "Download failed for '$t': $($_.Exception.Message)"
      }
    }
  }

  $after = @(Get-ChildItem -LiteralPath $msuRoot -Recurse -File -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName)
  $downloaded = @($after | Where-Object { $before -notcontains $_ })

  return [pscustomobject]@{ Selected=$selected; Downloaded=$downloaded }
}

# ==============================
# Helper functions for KB/architecture extraction
# ==============================
function Get-KbNumberFromPath([string]$p) {
  $m = [regex]::Match($p, '(?i)\bkb(\d{6,8})\b')
  if ($m.Success) { return [int]$m.Groups[1].Value }
  return [int]::MaxValue
}

function Get-ArchFromPath([string]$p) {
  $leaf = Split-Path -Leaf $p
  $l = $leaf.ToLowerInvariant()
  
  if ($l -like '*arm64*') { return 'arm64' }
  if ($l -like '*x64*' -or $l -like '*amd64*') { return 'x64' }
  if ($l -like '*x86*') { return 'x86' }
  
  return $null
}

# ==============================
# DU folder preparation (organize by category and KB number)
# ==============================
function Initialize-DUFolders {
  param(
    [Parameter(Mandatory=$true)][string]$DuRoot,
    [Parameter(Mandatory=$true)][string]$IsoFolder
  )

  $msuRoot = Join-Path $IsoFolder $Config.MsusSubdirName

  # Validate that MSU files exist in the msus subdirectory
  $all = @(Get-ChildItem -LiteralPath $msuRoot -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in @('.msu','.cab') })
  if ($all.Count -lt 1) {
    Stop-Script "No MSU/CAB files found in msus folder: $msuRoot. MSUs must be present to proceed."
  }

  # Clean and recreate DU root (removes stale category/KB folders)
  Write-Verbose "Cleaning DU root structure: $DuRoot"
  if (Test-Path $DuRoot) {
    try { Remove-Item -Path $DuRoot -Recurse -Force -ErrorAction SilentlyContinue } catch {}
  }
  New-Folder $DuRoot

  # Process each category subdirectory under msus\
  # Returns a pscustomobject with one ordered hashtable per category:
  #   each hashtable maps "KB#####" -> folder path under $DuRoot\<category>\KB#####\
  $duFolders = [pscustomobject]@{
    SSU     = [ordered]@{}
    LCU     = [ordered]@{}
    SafeOS  = [ordered]@{}
    SetupDU = [ordered]@{}
  }

  $totalFiles = 0
  $totalKbs = 0
  foreach ($cat in @('SSU','LCU','SafeOS','SetupDU')) {
    $catSrcDir = Join-Path $msuRoot $cat
    if (-not (Test-Path $catSrcDir)) {
      Write-Verbose ("No {0} directory found under {1}; skipping." -f $cat, $msuRoot)
      continue
    }

    $catFiles = @(Get-ChildItem -LiteralPath $catSrcDir -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in @('.msu','.cab') })
    if ($catFiles.Count -lt 1) {
      Write-Verbose ("No MSU/CAB files found in {0}; skipping." -f $catSrcDir)
      continue
    }

    $catDuDir = Join-Path $DuRoot $cat
    New-Folder $catDuDir

    # Group files by KB number, filtering by detected architecture
    $kbMap = @{}
    foreach ($f in $catFiles) {
      $kb = Get-KbNumberFromPath -p $f.FullName
      if ($kb -eq [int]::MaxValue) {
        Write-Host ("WARNING: Could not extract KB number from {0}; skipping." -f $f.Name) -ForegroundColor Yellow
        continue
      }

      $fileArch = Get-ArchFromPath -p $f.FullName
      if ($fileArch -and $fileArch -ne $script:State.DetectedArch) {
        Write-Verbose ("Skipping {0} (arch {1} does not match detected {2})" -f $f.Name, $fileArch, $script:State.DetectedArch)
        continue
      }

      if (-not $kbMap.ContainsKey($kb)) { $kbMap[$kb] = @() }
      $kbMap[$kb] += $f.FullName
    }

    if ($kbMap.Count -lt 1) {
      Write-Host ("WARNING: Architecture filter removed all {0} files (detected: {1})" -f $cat, $script:State.DetectedArch) -ForegroundColor Yellow
      continue
    }

    # Create KB-numbered folders under DuRoot\<category>\ and copy files
    $sortedKbs = @($kbMap.Keys | Sort-Object { [int]$_ })
    foreach ($kb in $sortedKbs) {
      $kbKey = "KB$([int]$kb)"
      $kbFolder = Join-Path $catDuDir $kbKey
      New-Folder $kbFolder
      $duFolders.$cat[$kbKey] = $kbFolder
      foreach ($f in $kbMap[$kb]) {
        $fname = Split-Path -Leaf $f
        Copy-Item -Path $f -Destination (Join-Path $kbFolder $fname) -Force
        Write-Verbose ("  Copied {0} -> {1}\{2}" -f $fname, $cat, $kbKey)
        $totalFiles++
      }
      $totalKbs++
    }

    Write-Verbose ("Category {0}: {1} KB folder(s)" -f $cat, $duFolders.$cat.Count)
  }

  Write-Host ("Organized {0} MSU/CAB file(s) into {1} KB folder(s) across categories (SSU:{2}, LCU:{3}, SafeOS:{4}, SetupDU:{5})." -f `
    $totalFiles, $totalKbs, $duFolders.SSU.Count, $duFolders.LCU.Count, $duFolders.SafeOS.Count, $duFolders.SetupDU.Count) -ForegroundColor Cyan
  return $duFolders
}

# ==============================
# Package application helpers
# ==============================
function Add-PackagesOrdered {
  param(
    [Parameter(Mandatory=$true)][string]$MountDir,
    [Parameter(Mandatory=$true)][hashtable]$KbFolders,
    [Parameter(Mandatory=$true)][string]$ScratchRoot,
    [Parameter(Mandatory=$true)][string]$LogBasePath,
    [Parameter(Mandatory=$true)][string]$ContextLabel
  )

  if ($KbFolders.Count -lt 1) {
    Write-Host ("No KB folders provided for {0}; skipping." -f $ContextLabel) -ForegroundColor Yellow
    return
  }

  Write-Host ("Applying packages to {0} in KB order (total {1} KBs)..." -f $ContextLabel, $KbFolders.Count) -ForegroundColor Cyan

  $succeeded = @()
  $failed    = @()

  foreach ($kb in ($KbFolders.Keys | Sort-Object)) {
    Assert-NotCancelled
    $kbFolder = $KbFolders[$kb]
    $pkgFiles = @()
    $pkgFiles += @(Get-ChildItem -LiteralPath $kbFolder -Filter "*.msu" -File -ErrorAction SilentlyContinue | Sort-Object Name | Select-Object -ExpandProperty FullName)
    $pkgFiles += @(Get-ChildItem -LiteralPath $kbFolder -Filter "*.cab" -File -ErrorAction SilentlyContinue | Sort-Object Name | Select-Object -ExpandProperty FullName)

    foreach ($pkg in $pkgFiles) {
      Assert-NotCancelled
      $leaf   = Split-Path $pkg -Leaf
      $pkgLog = $LogBasePath.Replace(".log", ("_KB{0}_{1}.log" -f $kb, (Protect-Token $leaf)))
      $label  = ("Add-Package {0} {1}: {2}" -f $ContextLabel, $kb, $leaf)

      $rc = Invoke-External -FilePath $script:State.DismPath -ArgumentList @(
        "/Image:$MountDir",
        "/Add-Package",
        "/PackagePath:$pkg",
        "/ScratchDir:$ScratchRoot",
        "/LogPath:$pkgLog"
      ) -StepName $label

      if ($rc -eq 0) {
        $succeeded += $pkg
      } else {
        Write-Host ("WARNING: Failed to add {0}/{1} (exit {2}); skipping. See log: {3}" -f $kb, $leaf, $rc, $pkgLog) -ForegroundColor Yellow
        $failed += $pkg
      }
    }
  }

  $total = $succeeded.Count + $failed.Count
  if ($total -eq 0) {
    Write-Host ("No packages found for {0}; skipping." -f $ContextLabel) -ForegroundColor Yellow
    return
  }

  Write-Host ""
  Write-Host ("Package apply summary for {0}:" -f $ContextLabel) -ForegroundColor Cyan
  Write-Host ("  Succeeded: {0}" -f $succeeded.Count) -ForegroundColor Green
  foreach ($s in $succeeded) { Write-Verbose ("    OK: {0}" -f $s) }
  if ($failed.Count -gt 0) {
    Write-Host ("  Failed:    {0}" -f $failed.Count) -ForegroundColor Yellow
    foreach ($f in $failed) { Write-Host ("    FAILED: {0}" -f $f) -ForegroundColor Yellow }
  }
  Write-Host ""
}

function Add-PackagesToISO {
  param(
    [Parameter(Mandatory=$true)][string]$IsoRoot,
    [Parameter(Mandatory=$true)][hashtable]$KbFolders,
    [Parameter(Mandatory=$true)][string]$ContextLabel
  )

  if ($KbFolders.Count -lt 1) {
    Write-Host ("No KB folders provided for {0}; skipping." -f $ContextLabel) -ForegroundColor Yellow
    return
  }

  Write-Host ("Writing packages for {0} to ISO..." -f $ContextLabel) -ForegroundColor Cyan

  $succeeded = @()
  $failed    = @()

  foreach ($kb in ($KbFolders.Keys | Sort-Object)) {
    Assert-NotCancelled
    $kbFolder = $KbFolders[$kb]
    $pkgFiles = @()
    $pkgFiles += @(Get-ChildItem -LiteralPath $kbFolder -Filter "*.msu" -File -ErrorAction SilentlyContinue | Sort-Object Name | Select-Object -ExpandProperty FullName)
    $pkgFiles += @(Get-ChildItem -LiteralPath $kbFolder -Filter "*.cab" -File -ErrorAction SilentlyContinue | Sort-Object Name | Select-Object -ExpandProperty FullName)

    foreach ($pkg in $pkgFiles) {
      Assert-NotCancelled
      $leaf = Split-Path $pkg -Leaf
      Write-Host ("Copying package {0}: {1}" -f $kb, $leaf) -ForegroundColor Cyan
      $dst = Join-Path $IsoRoot $leaf
      Copy-Item -Path $pkg -Destination $dst -Force
      if ($?) {
        $succeeded += $pkg
      } else {
        Write-Host ("WARNING: Failed to add {0}" -f $leaf) -ForegroundColor Yellow
        $failed += $pkg
      }
    }
  }

  $total = $succeeded.Count + $failed.Count
  if ($total -eq 0) {
    Write-Host ("No packages found for {0}; skipping." -f $ContextLabel) -ForegroundColor Yellow
    return
  }

  Write-Host ""
  Write-Host ("Package apply summary for {0}:" -f $ContextLabel) -ForegroundColor Cyan
  Write-Host ("  Succeeded: {0}" -f $succeeded.Count) -ForegroundColor Green
  foreach ($s in $succeeded) { Write-Verbose ("    OK: {0}" -f $s) }
  if ($failed.Count -gt 0) {
    Write-Host ("  Failed:    {0}" -f $failed.Count) -ForegroundColor Yellow
    foreach ($f in $failed) { Write-Host ("    FAILED: {0}" -f $f) -ForegroundColor Yellow }
  }
  Write-Host ""
}

# ==============================
# WinRE-in-OS servicing helper
# ==============================
function Update-WinREInsideMountedOS {
  param(
    [Parameter(Mandatory=$true)][string]$OsMountDir,
    [Parameter(Mandatory=$true)][string]$OsName,
    [Parameter(Mandatory=$true)][int]$OsIndex,
    [Parameter(Mandatory=$true)][object]$DuFolders,
    [Parameter(Mandatory=$true)][string]$ScratchRoot,
    [Parameter(Mandatory=$true)][string]$LogsRoot
  )

  $winrePath = Join-Path $OsMountDir 'Windows\System32\Recovery\winre.wim'
  if (-not (Test-Path $winrePath)) {
    Write-Verbose ("WinRE not found for OS: {0} (index {1}); skipping WinRE servicing." -f $OsName, $OsIndex)
    return
  }

  $nameTag = Protect-Token $OsName
  $tmp = Join-Path $script:State.MountRoot ("winre_{0}_idx{1}" -f $nameTag, $OsIndex)
  $tmpWim = Join-Path $tmp "winre.wim"
  $mDir  = Join-Path $tmp "MOUNT"
  New-Folder $tmp
  New-Folder $mDir

  Write-Verbose ("Copy winre.wim for OS: {0} (index {1})" -f $OsName, $OsIndex)
  Invoke-Step ("Copy winre.wim for OS: $OsName (index $OsIndex)") { Copy-Item -Path $winrePath -Destination $tmpWim -Force } | Out-Null

  Assert-ImageNotMounted -ImageFile $tmpWim -Index 1

  $mountLog = Join-Path $LogsRoot ("dism_mount_winre_{0}_idx{1}.log" -f $nameTag, $OsIndex)
  Write-Verbose ("Mount WinRE for {0} (index {1})" -f $OsName, $OsIndex)
  $rc = Invoke-External -FilePath $script:State.DismPath -ArgumentList @(
    "/Mount-Image",
    "/ImageFile:$tmpWim",
    "/Index:1",
    "/MountDir:$mDir",
    "/ScratchDir:$ScratchRoot",
    "/LogPath:$mountLog"
  ) -StepName ("Mount WinRE for {0} (index {1})" -f $OsName, $OsIndex)

  if ($rc -ne 0) {
    Write-Host ("WARNING: Failed to mount WinRE for {0} (index {1}). Skipping WinRE updates." -f $OsName, $OsIndex) -ForegroundColor Yellow
    return
  }

  try {
    # WinRE targets: SafeOS DU
    if ($DuFolders.SafeOS.Count -gt 0) {
      $safeOsBase = Join-Path $LogsRoot ("dism_addpackage_winre_safeos_{0}_idx{1}.log" -f $nameTag, $OsIndex)
      Add-PackagesOrdered -MountDir $mDir -KbFolders $DuFolders.SafeOS -ScratchRoot $ScratchRoot -LogBasePath $safeOsBase -ContextLabel ("WinRE for {0} SafeOS" -f $OsName)
    }
  }
  finally {
    Write-Verbose ("Unmount and commit WinRE for {0} (index {1})" -f $OsName, $OsIndex)
    $rc2 = Invoke-External -FilePath $script:State.DismPath -ArgumentList @(
      "/Unmount-Image",
      "/MountDir:$mDir",
      "/Commit"
    ) -StepName ("Commit WinRE for {0} (index {1})" -f $OsName, $OsIndex)

    if ($rc2 -ne 0) { Write-Host ("WARNING: Failed to unmount/commit WinRE for {0} (index {1})." -f $OsName, $OsIndex) -ForegroundColor Yellow }
  }

  Invoke-Step ("Replace WinRE in install.wim image ({0})" -f $OsName) { Copy-Item -Path $tmpWim -Destination $winrePath -Force } | Out-Null
}

# ==============================
# WIM servicing (install.wim, boot.wim)
# ==============================
function Update-InstallWimIndex {
  param(
    [string]$WimPath,
    [int]$Index,
    [string]$IndexName,
    [object]$DuFolders,
    [string]$MountRoot,
    [string]$LogsRoot,
    [string]$ScratchRoot
  )

  $nameTag = Protect-Token $IndexName

  Write-Host ("Servicing install.wim: {0} (index {1})" -f $IndexName, $Index) -ForegroundColor Cyan

  $mountDir = Join-Path $MountRoot ("os_{0}_idx{1}" -f $nameTag, $Index)
  $mountLog = Join-Path $LogsRoot ("dism_mount_os_{0}_idx{1}.log" -f $nameTag, $Index)

  Invoke-Step ("Prepare mount dir for install.wim: $IndexName (index $Index)") {
    if (Test-Path $mountDir) { Remove-Item $mountDir -Recurse -Force -ErrorAction SilentlyContinue | Out-Null }
    New-Item -ItemType Directory -Path $mountDir -Force | Out-Null
  } | Out-Null

  Assert-ImageNotMounted -ImageFile $WimPath -Index $Index

  Write-Verbose ("Mount install.wim image: {0} (index {1})" -f $IndexName, $Index)
  $rc = Invoke-External -FilePath $script:State.DismPath -ArgumentList @(
    "/Mount-Image",
    "/ImageFile:$WimPath",
    "/Index:$Index",
    "/MountDir:$mountDir",
    "/ScratchDir:$ScratchRoot",
    "/LogPath:$mountLog"
  ) -StepName ("Mount install.wim {0} (index {1})" -f $IndexName, $Index)

  if ($rc -ne 0) {
    Write-Host ("WARNING: Failed to mount install.wim {0} (index {1}). Skipping." -f $IndexName, $Index) -ForegroundColor Yellow
    return
  }

  try {
    # install.wim targets: SSU (prerequisites) -> LCU (checkpoint chain, KB order)
    if ($DuFolders.SSU.Count -gt 0) {
      $ssuBase = Join-Path $LogsRoot ("dism_addpackage_install_wim_{0}_idx{1}_ssu.log" -f $nameTag, $Index)
      Add-PackagesOrdered -MountDir $mountDir -KbFolders $DuFolders.SSU -ScratchRoot $ScratchRoot -LogBasePath $ssuBase -ContextLabel ("install.wim for {0} SSU" -f $IndexName)
    }
    if ($DuFolders.LCU.Count -gt 0) {
      $lcuBase = Join-Path $LogsRoot ("dism_addpackage_install_wim_{0}_idx{1}_lcu.log" -f $nameTag, $Index)
      Add-PackagesOrdered -MountDir $mountDir -KbFolders $DuFolders.LCU -ScratchRoot $ScratchRoot -LogBasePath $ssuBase -ContextLabel ("install.wim for {0} LCU" -f $IndexName)
    }

    # Service WinRE inside this install.wim image (if it exists)
    Update-WinREInsideMountedOS -OsMountDir $mountDir -OsName $IndexName -OsIndex $Index -DuFolders $DuFolders -ScratchRoot $ScratchRoot -LogsRoot $LogsRoot
  }
  finally {
    Write-Verbose ("Unmount and commit install.wim image: {0} (index {1})" -f $IndexName, $Index)
    $rc2 = Invoke-External -FilePath $script:State.DismPath -ArgumentList @(
      "/Unmount-Image",
      "/MountDir:$mountDir",
      "/Commit"
    ) -StepName ("Commit install.wim {0} (index {1})" -f $IndexName, $Index)

    if ($rc2 -ne 0) { Write-Host ("WARNING: Failed to unmount/commit install.wim {0} (index {1})." -f $IndexName, $Index) -ForegroundColor Yellow }
  }
}

function Update-BootWim {
  param(
    [Parameter(Mandatory=$true)][string]$BootWimPath,
    [Parameter(Mandatory=$true)][object]$DuFolders,
    [string]$MountRoot,
    [string]$LogsRoot,
    [string]$ScratchRoot
  )

  # Boot.wim contains WinPE indices used for Setup/deployment
  $pairs = Get-WimPairs -InstallPath $BootWimPath
  $idxs = @($pairs | Select-Object -ExpandProperty Index)

  foreach ($idx in $idxs) {
    Assert-NotCancelled
    $nm = ($pairs | Where-Object { $_.Index -eq $idx } | Select-Object -First 1).Name
    if (-not $nm) { $nm = "<unknown>" }
    $nameTag = Protect-Token $nm

    Write-Host ("Servicing boot.wim: {0} (index {1})" -f $nm, $idx) -ForegroundColor Cyan

    $mountDir = Join-Path $MountRoot ("boot_{0}_idx{1}" -f $nameTag, $idx)
    $mountLog = Join-Path $LogsRoot ("dism_mount_boot_{0}_idx{1}.log" -f $nameTag, $idx)

    Invoke-Step ("Prepare mount dir for boot.wim: $nm (index $idx)") {
      if (Test-Path $mountDir) { Remove-Item $mountDir -Recurse -Force -ErrorAction SilentlyContinue | Out-Null }
      New-Item -ItemType Directory -Path $mountDir -Force | Out-Null
    } | Out-Null

    Assert-ImageNotMounted -ImageFile $BootWimPath -Index $idx

    Write-Verbose ("Mount boot.wim: {0} (index {1})" -f $nm, $idx)
    $rc = Invoke-External -FilePath $script:State.DismPath -ArgumentList @(
      "/Mount-Image",
      "/ImageFile:$BootWimPath",
      "/Index:$idx",
      "/MountDir:$mountDir",
      "/ScratchDir:$ScratchRoot",
      "/LogPath:$mountLog"
    ) -StepName ("Mount boot.wim {0} (index {1})" -f $nm, $idx)

    if ($rc -ne 0) {
      Write-Host ("WARNING: Failed to mount boot.wim {0} (index {1}). Skipping." -f $nm, $idx) -ForegroundColor Yellow
      continue
    }

    try {
      # boot.wim targets: SafeOS DU (WinRE/recovery)
      if ($DuFolders.SafeOS.Count -gt 0) {
        $safeOsBase = Join-Path $LogsRoot ("dism_addpackage_boot_safeos_{0}_idx{1}.log" -f $nameTag, $idx)
        Add-PackagesOrdered -MountDir $mountDir -KbFolders $DuFolders.SafeOS -ScratchRoot $ScratchRoot -LogBasePath $safeOsBase -ContextLabel ("boot.wim for {0} SafeOS" -f $nm)
      }
    }
    finally {
      Write-Verbose ("Unmount and commit boot.wim: {0} (index {1})" -f $nm, $idx)
      $rc2 = Invoke-External -FilePath $script:State.DismPath -ArgumentList @(
        "/Unmount-Image",
        "/MountDir:$mountDir",
        "/Commit"
      ) -StepName ("Commit boot.wim {0} (index {1})" -f $nm, $idx)

      if ($rc2 -ne 0) { Write-Host ("WARNING: Failed to unmount/commit boot.wim {0} (index {1})." -f $nm, $idx) -ForegroundColor Yellow }
    }
  }
}

# ==============================
# Build ISO
# ==============================
function Build-Iso([string]$IsoRoot, [string]$OutputIso) {
  $etfs = Join-Path $IsoRoot $Config.BootFileBIOS
  $efis = Join-Path $IsoRoot $Config.BootFileUEFI
  if (-not (Test-Path $etfs)) { Stop-Script "Missing BIOS boot file: $etfs" }
  if (-not (Test-Path $efis)) { Stop-Script "Missing UEFI boot file: $efis" }

  $bootdata = "2#p0,e,b$etfs#pEF,e,b$efis"
  $oscdimgfsargs = @() + $Config.OscdimgFsArgs + @("-l$($Config.IsoVolumeLabel)", "-bootdata:$bootdata", $IsoRoot, $OutputIso)

  Write-Host "Building ISO: $OutputIso" -ForegroundColor Green
  $rc = Invoke-External -FilePath $script:State.OscdimgPath -ArgumentList $oscdimgfsargs -StepName "OSCDIMG Build ISO"
  if ($rc -ne 0) { Stop-Script "oscdimg failed with exit code $rc" }
}

# ==============================
# Cleanup
# ==============================
function Clear-Hardened {
  param(
    [switch]$Aggressive,
    [switch]$FromCancel
  )

  try { Stop-TrackedChildren } catch {}

  if ($Aggressive) {
    try { Stop-DismProcessesForWorkRoot -WorkRoot $script:State.WorkRoot } catch {}
    try { Stop-Process -Name "dism" -Force -ErrorAction SilentlyContinue } catch {}  # catch any stray DISM processes not tracked by WorkRoot
  }

  try {
    if ($script:State.MountRoot) {
      Clear-PreflightDismMounts -RelevantImageFiles @() -RelevantMountRoot $script:State.MountRoot
    } else {
      Clear-DismMountPoints
    }
  } catch {}

  try {
    if ($script:State.IsoWasMounted -and $script:State.IsoPath) {
      Dismount-DiskImage -ImagePath $script:State.IsoPath -ErrorAction SilentlyContinue | Out-Null
    }
  } catch {}

  if ($FromCancel) {
    try { Write-Host "Cleanup complete. Exiting." -ForegroundColor Yellow } catch {}
  }
}

# ==============================
# Argument parsing
# ==============================
$FolderArg = $null
$IsoPath = $null
$DestIsoPath = $null
$UseSystemTemp = $false
$DebugSwitch = $false
$VerboseSwitch = $false
$UseADK = $false
$UseSystem = $false
$ExplicitDism = $null
$ExplicitOscdimg = $null
$CleanWork = $false
$UpdateISO = $false
$UpdateMSUs = $false
$OnlyLatestLCU = $false
$CleanMSUs = $false
$ShowIndices = $false
$SelectHome = $false
$SelectPro = $false
$IndicesSpec = $null

for ($i = 0; $i -lt $args.Count; $i++) {
  $a = $args[$i]
  switch -Regex ($a) {
    '^(?:-h|-help|-\?|/\?)$'   { Show-Usage; exit 0 }
    '^(?:-UseSystemTemp)$'     { $UseSystemTemp = $true; continue }
    '^(?:-Verbose|-v)$'        { $VerboseSwitch = $true; continue }
    '^(?:-Debug|-d)$'          { $DebugSwitch = $true; continue }
    '^(?:-DryRun)$'            { $script:DryRun = $true; continue }
    '^(?:-UseADK)$'            { $UseADK = $true; continue }
    '^(?:-UseSystem)$'         { $UseSystem = $true; continue }
    '^(?:-CleanWork)$'         { $CleanWork = $true; continue }
    '^(?:-UpdateISO)$'         { $UpdateISO = $true; continue }
    '^(?:-UpdateMSUs)$'        { $UpdateMSUs = $true; continue }
    '^(?:-CleanMSUs)$'         { $CleanMSUs = $true; continue }
    '^(?:-OnlyLatestLCU)$'     { $OnlyLatestLCU = $true; continue }
    '^(?:-ShowIndices)$'       { $ShowIndices = $true; continue }
    '^(?:-Show)$'              { $ShowIndices = $true; continue }
    '^(?:-Home)$'              { $SelectHome = $true; continue }
    '^(?:-Pro)$'               { $SelectPro = $true; continue }
    '^(?:-Indices)$' {
      if ($i + 1 -ge $args.Count) { Stop-Script "-Indices requires one or more selector tokens" }
      $vals = New-Object System.Collections.Generic.List[string]
      $j = $i + 1
      while ($j -lt $args.Count) {
        $n = $args[$j]
        if ($n -is [string] -and $n.StartsWith('-')) { break }
        if ($n -is [System.Array] -and -not ($n -is [string])) { foreach ($e in $n) { $vals.Add([string]$e) | Out-Null } }
        else { $vals.Add([string]$n) | Out-Null }
        $j++
      }
      $IndicesSpec = ($vals -join ',')
      $i = $j - 1
      continue
    }
    '^(?:-dism)$'    { $ExplicitDism = $args[++$i]; continue }
    '^(?:-oscdimg)$' { $ExplicitOscdimg = $args[++$i]; continue }
    '^(?:-ISO)$'     { $IsoPath = $args[++$i]; continue }
    '^(?:-SrcISO)$'  { $IsoPath = $args[++$i]; continue }
    '^(?:-DestISO)$' { $DestIsoPath = $args[++$i]; continue }
    '^-{1,2}.*'      { Stop-Script "Unknown switch: $a" }
    default {
      if ($FolderArg) { Stop-Script "Only one folder argument allowed. Extra value: $a" }
      $FolderArg = $a
    }
  }
}

if ($DebugSwitch) { $DebugPreference = 'Continue' }
if ($VerboseSwitch) { $VerbosePreference = 'Continue' }
if (-not $FolderArg) { $FolderArg = (Get-Location).Path }

# ==============================
# MAIN
# ==============================
Register-CancelHandler
Assert-Admin
Resolve-Tools -ExplicitDism $ExplicitDism -ExplicitOscdimg $ExplicitOscdimg -UseADK:($UseADK) -UseSystem:($UseSystem)

try {
  $folderPath = (Resolve-Path -LiteralPath $FolderArg -ErrorAction SilentlyContinue).Path
  if (-not $folderPath) { Stop-Script "Folder not found: $FolderArg" }

  # Resolve ISO path (new run) or meta (update run)
  if (-not $UpdateISO) {
    if ($IsoPath) {
      $script:State.IsoPath = (Resolve-Path -LiteralPath $IsoPath -ErrorAction SilentlyContinue).Path
      if (-not $script:State.IsoPath) { Stop-Script "ISO file not found: $IsoPath" }
      $isoDir = Split-Path -Parent $script:State.IsoPath
    } else {
      $script:State.IsoPath = Get-InputIso -FolderPath $folderPath
      $isoDir = $folderPath
    }
    Initialize-WorkPaths -FolderPath $folderPath -IsoPath $script:State.IsoPath -UseSystemTemp:($UseSystemTemp)
  } else {
    $workBase = if ($UseSystemTemp) { [System.IO.Path]::GetTempPath() } else { Join-Path $folderPath $Config.WorkParentSubfolder }
    $metaFiles = Get-ChildItem -LiteralPath $workBase -Filter $Config.MetaFileName -Recurse -File -ErrorAction SilentlyContinue
    if ($metaFiles.Count -ne 1) { Stop-Script "UpdateISO requires exactly one meta file under $workBase. Found $($metaFiles.Count)." }
    $meta = Read-Meta -MetaPath $metaFiles[0].FullName
    if (-not $meta) { Stop-Script "Failed to read meta file: $($metaFiles[0].FullName)" }

    $script:State.IsoPath = $meta["ISOPath"]
    $script:State.IsoBaseName = $meta["IsoBaseName"]
    $script:State.WorkRoot = $meta["WorkRoot"]
    $script:State.IsoRoot = $meta["IsoRoot"]
    $script:State.InstallRoot = $meta["InstallRoot"]
    $script:State.MountRoot = $meta["MountRoot"]
    $script:State.LogsRoot = $meta["LogsRoot"]
    $script:State.ScratchRoot = $meta["ScratchRoot"]
    $script:State.DuRoot = $meta["DuRoot"]
    $script:State.StashedInstallWim = $meta["StashedInstallWim"]
    $script:State.OutputIsoPath = $meta["OutputIsoPath"]
    $script:State.MetaPath = $metaFiles[0].FullName
    $isoDir = Split-Path -Parent $script:State.IsoPath
  }

  if ($DestIsoPath) {
    $destDir = Split-Path -Parent $DestIsoPath
    $destDir = (Resolve-Path -LiteralPath $destDir -ErrorAction SilentlyContinue).Path
    if (-not $destDir) { Stop-Script "Destination directory not found: $(Split-Path -Parent $DestIsoPath)" }
    $script:State.OutputIsoPath = Join-Path $destDir ([IO.Path]::GetFileName($DestIsoPath))
  } elseif (-not $script:State.OutputIsoPath) {
    $baseName = [IO.Path]::GetFileNameWithoutExtension($script:State.IsoPath)
    if ($IsoPath) {
      $script:State.OutputIsoPath = Join-Path $isoDir ($baseName + $Config.OutputIsoSuffix)
    } else {
      $script:State.OutputIsoPath = Join-Path $folderPath ($baseName + $Config.OutputIsoSuffix)
    }
  }

  if ($CleanWork -and (Test-Path $script:State.WorkRoot)) {
    Remove-Item -Path $script:State.WorkRoot -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
  }

  # PREP always runs
  $script:DryRunPhase = 'prep'
  New-Folder $script:State.WorkRoot
  New-Folder $script:State.IsoRoot
  New-Folder $script:State.InstallRoot
  New-Folder $script:State.MountRoot
  New-Folder $script:State.LogsRoot
  New-Folder $script:State.ScratchRoot
  New-Folder $script:State.DuRoot

  # ISO extract (only if not UpdateISO)
  if (-not $UpdateISO) {
    $needsCopy = -not (Get-ChildItem -LiteralPath $script:State.IsoRoot -Force -ErrorAction SilentlyContinue | Select-Object -First 1)
    if ($needsCopy) {
      $isoMount = Mount-Iso $script:State.IsoPath
      Copy-IsoContents -SrcDrive $isoMount.Drive -DstFolder $script:State.IsoRoot -IsoPathForDisplay $script:State.IsoPath
    } else {
      Write-Host "ISO contents already extracted; skipping robocopy." -ForegroundColor Yellow
    }
  } else {
    Write-Host "UpdateISO: Reusing existing extracted ISO tree." -ForegroundColor Yellow
  }

  # Ensure extracted ISO tree is writable (fixes boot.wim mount/modify errors)
  Clear-ReadOnlyAttributes -Path $script:State.IsoRoot

  Assert-NotCancelled

  # Ensure $WinpeDriver$ exists in ISO root
  $driverFolderOnIso = Join-Path $script:State.IsoRoot $Config.DriverFolderName
  New-Folder $driverFolderOnIso
  $driverReadme = Join-Path $driverFolderOnIso "README.txt"
  if (-not (Test-Path $driverReadme)) {
    Set-Content -Path $driverReadme -Encoding ASCII -Value @(
      "Place INF-based drivers under this folder (subfolders allowed).",
      "Folder name is special: Windows Setup scans `"$($Config.DriverFolderName)`" at the root of install media for drivers."
    )
  }

  $isoSources = Join-Path $script:State.IsoRoot 'sources'
  if (-not (Test-Path $isoSources)) { Stop-Script "ISO sources directory not found: $isoSources" }

  # Write SetupConfig + launchers into ISO root
  Write-ConfigFiles -IsoRoot $script:State.IsoRoot
  Write-Cmds -IsoRoot $script:State.IsoRoot

  # Move install.* to INSTALL if needed (stash original install image)
  $stashWim = Join-Path $script:State.InstallRoot 'install.wim'
  $stashEsd = Join-Path $script:State.InstallRoot 'install.esd'
  if (-not (Test-Path $stashWim) -and -not (Test-Path $stashEsd)) {
    $srcInstall = Get-InstallImagePathFromRoot -Root $script:State.IsoRoot
    if (-not $srcInstall) { Stop-Script "Could not find sources\install.wim or sources\install.esd in ISO tree." }
    Move-Item -Path $srcInstall -Destination (Join-Path $script:State.InstallRoot (Split-Path -Leaf $srcInstall)) -Force
  }

  if (-not $script:State.StashedInstallWim) {
    $script:State.StashedInstallWim = Initialize-StashedInstallWim -InstallRoot $script:State.InstallRoot
  }
  if (-not $script:State.StashedInstallWim) { Stop-Script "INSTALL stash does not contain install.wim or install.esd." }

  Get-MediaInfoFromInstallWim -InstallWim $script:State.StashedInstallWim
  Show-RunBanner

  # AFTER PREP
  $script:DryRunPhase = 'afterprep'

  if ($ShowIndices) { Show-Indices -InstallPath $script:State.StashedInstallWim; return }

  $pairs = Get-WimPairs -InstallPath $script:State.StashedInstallWim
  $maxIndex = ($pairs | Measure-Object -Property Index -Maximum).Maximum
  $nameMap = Get-IndexNameMap -Pairs $pairs

  Write-Verbose ("install.wim stash: {0}" -f $script:State.StashedInstallWim)
  Write-Verbose ("Found {0} WIM indices; max index = {1}" -f $pairs.Count, $maxIndex)
  if ($VerbosePreference -eq 'Continue') {
    Write-Host "Available indices (index : name):" -ForegroundColor Cyan
    foreach ($p in $pairs) { Write-Host ("  {0}: {1}" -f $p.Index, $p.Name) -ForegroundColor Cyan }
    Write-Host ""
  }

  $explicitSelectionUsed = ($SelectHome -or $SelectPro -or ($IndicesSpec -and ([string]$IndicesSpec).Trim() -ne ""))

  $selected = @()
  $doRebuild = $true
  $doService = $true
  $doDownload = $true

  if ($UpdateISO -and -not $explicitSelectionUsed) {
    $selected = @()
    $doRebuild = $false
    $doService = $false
    $doDownload = $false
    Write-Host "[UpdateISO] No indices specified; skipping rebuild, servicing, and update handling." -ForegroundColor Yellow
  } else {
    if (-not $explicitSelectionUsed -and -not $UpdateISO) {
      $selected = 1..$maxIndex
    } else {
      if ($SelectHome) { $selected += (Resolve-LabelTokenToIndices -Pairs $pairs -Token "Home") }
      if ($SelectPro)  { $selected += (Resolve-LabelTokenToIndices -Pairs $pairs -Token "Pro") }

      if ($IndicesSpec -and ([string]$IndicesSpec).Trim() -ne "") {
        $parsed = ConvertFrom-IndicesSpec -Spec ([string]$IndicesSpec) -MaxIndex $maxIndex -Pairs $pairs
        if ($parsed.InvalidTokens.Count -gt 0 -or $parsed.InvalidIndices.Count -gt 0 -or $parsed.InvalidLabels.Count -gt 0) {
          Write-Host "Invalid selection detected:" -ForegroundColor Red
          if ($parsed.InvalidTokens.Count -gt 0)  { Write-Host ("  Invalid tokens: " + ($parsed.InvalidTokens -join ", ")) -ForegroundColor Red }
          if ($parsed.InvalidIndices.Count -gt 0) { Write-Host ("  Invalid indices: " + ($parsed.InvalidIndices -join ", ")) -ForegroundColor Red }
          if ($parsed.InvalidLabels.Count -gt 0)  { Write-Host ("  Unmatched labels/patterns: " + ($parsed.InvalidLabels -join ", ")) -ForegroundColor Red }
          Write-Host ""
          Write-Host "Available indices:" -ForegroundColor Cyan
          Show-Indices -InstallPath $script:State.StashedInstallWim
          exit 2
        }
        $selected += $parsed.Selected
      }

      $selected = @($selected | Sort-Object -Unique)
      if ($selected.Count -lt 1) { Stop-Script "Selection resulted in an empty set." }
    }
  }

  if ($selected.Count -gt 0) {
    Write-Host "Selected source indices (index : name):" -ForegroundColor Cyan
    foreach ($idx in $selected) {
      $nm = $nameMap[[int]$idx]; if (-not $nm) { $nm = "<unknown>" }
      Write-Host ("  {0}: {1}" -f $idx, $nm) -ForegroundColor Cyan
    }
  }

  if ($script:DryRun) {
    Write-Host "[DryRun] Would proceed with rebuild/download/servicing/ISO build." -ForegroundColor Yellow
    return
  }

  if (-not $doService) {
    Build-Iso -IsoRoot $script:State.IsoRoot -OutputIso $script:State.OutputIsoPath
    Write-Meta
    Write-Host "SUCCESS: Created bundled ISO:" -ForegroundColor Green
    Write-Host ("  " + $script:State.OutputIsoPath) -ForegroundColor Green
    Write-Host "Work folder preserved at:" -ForegroundColor Yellow
    Write-Host ("  " + $script:State.WorkRoot) -ForegroundColor Yellow
    return
  }

  # Rebuild ISO\sources\install.wim if requested
  if ($doRebuild) {
    $wimToService = Build-InstallWimFromSelection -SourceFile $script:State.StashedInstallWim -IsoSourcesDir $isoSources -SelectedSourceIndices $selected -MaxSourceIndex $maxIndex
  } else {
    $wimToService = Join-Path $isoSources 'install.wim'
    if (-not (Test-Path $wimToService)) {
      Stop-Script "Skipped rebuild, but ISO\sources\install.wim is missing. Specify indices to force rebuild."
    }
  }

  # Update cache semantics
  if ($CleanMSUs) { Clear-MsuFolder -FolderPath $isoDir }

  if ($doDownload) {
    $force = [bool]$UpdateMSUs -or [bool]$CleanMSUs
    if ($force) { Write-Host "[UpdateMSUs/CleanMSUs] Forcing update download/refresh..." -ForegroundColor Yellow }
    else { Write-Host "Ensuring updates exist in msus directory (download if missing)..." -ForegroundColor Yellow }

    if (-not $script:State.DetectedVersion) {
      Stop-Script "Could not detect OS version from ISO. Cannot query update catalog."
    }

    $catalogInfo = Initialize-AllMSUsPresent -IsoFolder $isoDir -OsName $script:State.DetectedOS -OsVersion $script:State.DetectedVersion -OsBuild $script:State.DetectedBuild -Arch $script:State.DetectedArch -ForceDownload $force
  }

  $msuRootPath = Join-Path $isoDir $Config.MsusSubdirName
  $localUpdates = @(Get-ChildItem -LiteralPath $msuRootPath -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in @('.msu','.cab') })
  if ($localUpdates.Count -gt 0) {
    # Organize MSU/CAB files from msus\<category>\ into DuRoot\<category>\KB####\ folders.
    # Returns a pscustomobject with per-category ordered KB folder maps.
    $duFolders = Initialize-DUFolders -DuRoot $script:State.DuRoot -IsoFolder $isoDir

    # Preflight: cleanup DISM mount state before servicing
    $bootWim = Join-Path $isoSources 'boot.wim'
    $relevant = @($wimToService)
    if (Test-Path $bootWim) { $relevant += $bootWim }
    Clear-PreflightDismMounts -RelevantImageFiles $relevant -RelevantMountRoot $script:State.MountRoot

    # Service install.wim indices: SSU -> LCU (per index); WinRE inside each index: SafeOS
    Write-Verbose "Starting install.wim servicing"
    $rebuiltPairs = Get-WimPairs -InstallPath $wimToService
    $rebuiltNameMap = Get-IndexNameMap -Pairs $rebuiltPairs
    $serviceIndexes = @($rebuiltPairs | Select-Object -ExpandProperty Index)

    Write-Host ("Servicing {0} install.wim index(es)..." -f $serviceIndexes.Count) -ForegroundColor Cyan
    foreach ($idx in $serviceIndexes) {
      Assert-NotCancelled
      $nm = $rebuiltNameMap[[int]$idx]; if (-not $nm) { $nm = "<unknown>" }
      Update-InstallWimIndex -WimPath $wimToService -Index $idx -IndexName $nm -DuFolders $duFolders -MountRoot $script:State.MountRoot -LogsRoot $script:State.LogsRoot -ScratchRoot $script:State.ScratchRoot
    }

    # Service boot.wim (WinPE/Setup): SafeOS DU
    if (Test-Path $bootWim) {
      Write-Host "Servicing boot.wim (WinPE)..." -ForegroundColor Cyan
      Update-BootWim -BootWimPath $bootWim -DuFolders $duFolders -MountRoot $script:State.MountRoot -LogsRoot $script:State.LogsRoot -ScratchRoot $script:State.ScratchRoot
    } else {
      Write-Host "WARNING: ISO\sources\boot.wim not found; skipping WinPE/Setup servicing." -ForegroundColor Yellow
    }

    # Finally copy SetupDU files into the root ISO tree (if any), so they're included in the ISO
    if ($duFolders.SetupDU.Count -gt 0) {
      Add-PackagesToISO -IsoRoot $script:State.IsoRoot -KbFolders $duFolders.SetupDU -ContextLabel "SetupDU files"
    }
  }

  # Build ISO
  Build-Iso -IsoRoot $script:State.IsoRoot -OutputIso $script:State.OutputIsoPath
  Write-Meta

  Write-Host "SUCCESS: Created bundled ISO:" -ForegroundColor Green
  Write-Host ("  " + $script:State.OutputIsoPath) -ForegroundColor Green
  Write-Host "Work folder preserved at:" -ForegroundColor Yellow
  Write-Host ("  " + $script:State.WorkRoot) -ForegroundColor Yellow
}
catch [System.OperationCanceledException] {
  if (-not $script:Cancelled) { Write-Host "Operation cancelled." -ForegroundColor Yellow }
}
finally {
  Clear-Hardened -Aggressive:($script:Cancelled) -FromCancel:($script:Cancelled)
}
