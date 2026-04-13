<#
.SYNOPSIS
Creates a bundled Windows installation ISO by extracting an input ISO, optionally rebuilding install.wim from selected editions, servicing the image(s), and generating a new bootable ISO.

.DESCRIPTION
This script operates on a folder that contains a Windows ISO (excluding *.bundled.iso) OR an explicitly provided ISO path.
It extracts the ISO into a stable work directory, prepares an installation image, optionally services it, writes an autounattend.xml,
and builds a new bootable ISO.

Help display:
- Use -h, -help, -? (or /?) to display this embedded help.
- The script displays help by calling Get-Help -Path <this script> -Full. Comment-based help is supported for scripts and functions.
- Use -Full to display all sections reliably, including NOTES.

Drivers folder:
- The script creates a folder named "Drivers" at the root of the extracted ISO work tree (same level as setup.exe) so it exists in the final ISO.
- autounattend.xml includes a windowsPE driver search path using Microsoft-Windows-PnpCustomizationsWinPE -> DriverPaths -> PathAndCredentials.
  The path used is: %configsetroot%\Drivers
  DriverPaths/PathAndCredentials is the standard container/list structure for specifying driver search paths in an answer file.

Architecture behavior:
- The script determines the ISO/image architecture (for example amd64 or arm64) from the installation image metadata and uses that architecture
  in the generated autounattend.xml component declarations where processorArchitecture is required.

Index selection:
- If no selection is provided, behavior depends on -UpdateISO:
  - Without -UpdateISO: defaults to ALL indices.
  - With -UpdateISO: defaults to EMPTY selection unless indices are explicitly specified.
- Explicit selection can be made using:
  - -Home, -Pro
  - -Indices with numbers, ranges, labels, wildcard labels (* and ?), or regex labels (re:<pattern>).

UpdateISO behavior:
- -UpdateISO reuses an existing work folder from a prior run (as identified by the meta file).
- If -UpdateISO is specified and NO explicit index selection is provided (-Home/-Pro/-Indices):
  - The script does not rebuild ISO\sources\install.wim
  - The script does not service the image(s)
  - The script does not merge/apply MSU packages
  - The script builds an ISO from the existing extracted ISO work tree.
- If -UpdateISO is specified and explicit indices are provided, rebuild/servicing/MSU handling occurs for the selected indices.

Sanity check (install image presence):
- When the script is in a mode that builds an ISO from an existing work tree (for example, -UpdateISO with no explicit indices),
  it verifies that ISO\sources\install.wim or ISO\sources\install.esd is present before building the final ISO. If missing, the script
  fails with a clear error to prevent creating a non-installable ISO.

MSU (offline update) behavior:
- When servicing is enabled (i.e., not the special -UpdateISO-with-no-indices skip mode), the script can apply Windows update packages
  (.msu files) to each mounted install.wim index using DISM /Add-Package with /PackagePath pointing to a folder. DISM supports using a
  folder as PackagePath and will process .msu/.cab packages discovered there. [1](https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/dism-operating-system-package-servicing-command-line-options?view=windows-11)
- The script treats the folder containing the source ISO as the MSU staging folder:
  - First, it searches for existing *.msu files in the source ISO folder (top-level).
  - If no MSUs are present, it automatically downloads the appropriate MSU package(s) for the Windows version/release and architecture
    represented by the ISO into that same source ISO folder, then uses that folder as the DISM PackagePath.
- Multiple MSUs may be required. For Windows 11 (notably version 24H2 and later), Microsoft can publish "checkpoint cumulative updates"
  that must be applied before a target cumulative update when sourcing updates from the Microsoft Update Catalog. [2](https://learn.microsoft.com/en-us/windows/deployment/update/catalog-checkpoint-cumulative-updates)
  In these cases, the script downloads all required prerequisite checkpoint MSU(s) and the latest cumulative MSU(s) for the ISO’s Windows
  release branch and architecture, placing them together in the source ISO folder for DISM folder-based installation. [2](https://learn.microsoft.com/en-us/windows/deployment/update/catalog-checkpoint-cumulative-updates)[1](https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/dism-operating-system-package-servicing-command-line-options?view=windows-11)
- The download source is the Microsoft Update Catalog, and selection is driven by the ISO’s detected Windows release (e.g., 24H2) and
  architecture (e.g., x64/arm64), so the chosen updates match the installation media. [3](https://www.catalog.update.microsoft.com/Search.aspx?q=24H2%2C%20cumulative)[4](https://www.catalog.update.microsoft.com/Search.aspx?q=-msu)
- -UpdateISO with no explicit indices continues to skip all rebuild/servicing/MSU handling (including auto-download). (See UpdateISO behavior above.)

DryRun behavior:
- With -DryRun, the script completes PREP actions needed to stage the work tree and then prints what would happen for post-PREP actions.

.PARAMETER Folder
Optional. Folder to process. If omitted, the current directory is used.
- When -ISO/-SrcISO is not used, this folder must contain exactly one input ISO (excluding *.bundled.iso).
- The folder containing the source ISO is also the default location where the script looks for *.msu files and, if needed,
  downloads the required MSU(s) for servicing.

.PARAMETER ISO
Optional. Explicit path to the source ISO file to use instead of searching the target folder for exactly one *.iso.
- This switch overrides the default ISO discovery behavior.
- When specified, the default destination ISO path (unless overridden by -DestISO) is derived from this ISO path.
- The folder containing this ISO becomes the MSU staging/download folder when servicing is enabled.

.PARAMETER SrcISO
Alias of -ISO. Provides an explicit path to the source ISO file.

.PARAMETER DestISO
Optional. Explicit path for the output ISO file.
- Overrides the default behavior of creating <base>.bundled.iso next to the input ISO (or next to the discovered ISO in the target folder).

.PARAMETER DryRun
Runs PREP actions, then prints what would happen for post-PREP actions.

.PARAMETER CleanWork
Deletes the stable work folder before starting.

.PARAMETER UpdateISO
Reuses an existing work folder. If used without explicit indices, no rebuild/servicing/MSU actions occur.

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

.EXAMPLE
PS> .\BundledWindowsIso.ps1 D:\temp -Pro -CleanWork
Builds a bundled ISO using the Pro edition only, deleting any prior work folder first.
The script discovers the source ISO in D:\temp (exactly one *.iso excluding *.bundled.iso must exist).

.EXAMPLE
PS> .\BundledWindowsIso.ps1 -ISO D:\Images\Win11_25H2_English_x64.iso -Pro
Builds a bundled ISO using the Pro edition only, using the explicitly provided ISO path.
The default destination ISO path is derived from the provided ISO path unless -DestISO is specified.

.EXAMPLE
PS> .\BundledWindowsIso.ps1 -ISO D:\Images\Win11.iso -DestISO D:\Out\Win11_Custom.iso -Indices "re:^Education( N)?$"
Builds a bundled ISO from a specific ISO, selecting only Education/Education N editions (regex label selection),
and writes the output to an explicit destination ISO path.

.EXAMPLE
PS> .\BundledWindowsIso.ps1 -ISO D:\Images\Win11_24H2_x64.iso -Pro
If no *.msu files are present in D:\Images, the script downloads the required MSU package(s) for that ISO’s Windows release/architecture
from the Microsoft Update Catalog into D:\Images and applies them during servicing. [3](https://www.catalog.update.microsoft.com/Search.aspx?q=24H2%2C%20cumulative)[2](https://learn.microsoft.com/en-us/windows/deployment/update/catalog-checkpoint-cumulative-updates)[1](https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/dism-operating-system-package-servicing-command-line-options?view=windows-11)

.EXAMPLE
PS> .\BundledWindowsIso.ps1 -UpdateISO
Reuses the existing work folder and performs no rebuild/servicing/MSU actions (no indices were specified).
The script verifies that ISO\sources\install.wim or install.esd exists before building the final ISO.

.EXAMPLE
PS> .\BundledWindowsIso.ps1 -UpdateISO -Pro
Reuses the existing work folder and forces rebuild/servicing using the Pro edition selection.

.EXAMPLE
PS> .\BundledWindowsIso.ps1 -Indices "* N*"
Selects all N editions (quote required due to space).

.EXAMPLE
PS> .\BundledWindowsIso.ps1 -h
Displays this help content.

.NOTES
- For help display, the script should resolve its own file path once (for example, $ScriptPath) and reuse it when calling
  Get-Help -Path $ScriptPath -Full.
- The Drivers folder is intended for INF-based drivers (subfolders allowed) and is referenced via the answer file driver path.
- DISM package servicing supports using /PackagePath pointing to a folder that contains multiple .msu/.cab packages. [1](https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/dism-operating-system-package-servicing-command-line-options?view=windows-11)
- Starting with Windows 11, version 24H2, Microsoft introduced checkpoint cumulative updates; when sourcing updates from the Microsoft Update Catalog,
  devices/images may need to take all prior checkpoint cumulative updates before applying a target update. [2](https://learn.microsoft.com/en-us/windows/deployment/update/catalog-checkpoint-cumulative-updates)
- The Microsoft Update Catalog exposes cumulative updates by Windows version and architecture (e.g., “Windows 11 Version 24H2 for x64-based Systems”). [3](https://www.catalog.update.microsoft.com/Search.aspx?q=24H2%2C%20cumulative)[4](https://www.catalog.update.microsoft.com/Search.aspx?q=-msu)

#>

# ==============================
# BundledWindowsIso.ps1
# ==============================

# ==============================
# Configuration (editable)
# ==============================
$Config = [ordered]@{
  # Output naming
  OutputIsoSuffix          = '.bundled.iso'

  # Work layout (note: WorkRoot is now just the ISO base name under this parent)
  WorkParentSubfolder      = '_WinIsoBundlerWork'
  WorkIsoSubdir            = 'ISO'
  WorkInstallSubdir        = 'INSTALL'
  WorkMountSubdir          = 'MOUNT'
  WorkLogsSubdir           = 'LOGS'
  WorkScratchSubdir        = 'SCRATCH'
  MetaFileName             = 'bundler.meta.txt'

  # Driver injection folder on ISO root
  DriverFolderName         = 'Drivers'
  DriverFolderUnattendPath = $null

  # Robocopy base args (always applied)
  RobocopyArgsBase         = @('/E','/R:2','/W:2','/DCOPY:DA','/COPY:DAT')
  # Quiet unless -Verbose
  RobocopyArgsQuiet        = @('/NFL','/NDL','/NJH','/NJS','/NC','/NS','/NP')

  # oscdimg ISO boot settings
  BootFileBIOS             = 'boot\etfsboot.com'
  BootFileUEFI             = 'efi\microsoft\boot\efisys.bin'
  IsoVolumeLabel           = 'WIN_BUNDLED'
  OscdimgFsArgs            = @('-m','-o','-u2','-udfver102')

  # First-logon winget script
  AutounattendName         = 'autounattend.xml'
  PostInstallScriptName    = 'PostInstallWinget.ps1'
  WingetLogFile            = 'WingetPostInstall.log'
  WingetArgs               = @('--all','--include-unknown','--silent','--accept-package-agreements','--accept-source-agreements')

  # MSU selection filtering
  ExcludeTitleTokens        = @('preview', '.net', 'dynamic', 'safe os', 'setup dynamic')
  ExcludeServerTitleTokens  = @('microsoft server operating system', 'windows server')

  DefaultWin10Release       = '22H2'
  DefaultWin11Release       = '25H2'

  # Checkpoint prerequisite(s) (best-effort) for Win11 24H2+ offline servicing workflows
  KnownCheckpointKBs24H2Plus = @('KB5043080')

  # ADK probing (user preference: only amd64 + arm64)
  AdkToolArch               = @('amd64','arm64')
  AdkDeploymentToolsRoot    = "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools"
}
$Config.DriverFolderUnattendPath = "%configsetroot%\$($Config.DriverFolderName)"

# ==============================
# Script state (do not edit)
# ==============================
$script:State = [ordered]@{
  IsoPath               = $null
  IsoWasMounted         = $false

  WorkBase              = $null
  WorkRoot              = $null
  IsoRoot               = $null
  InstallRoot           = $null
  MountRoot             = $null
  LogsRoot              = $null
  ScratchRoot           = $null
  MetaPath              = $null

  DismPath              = $null
  DismLabel             = $null
  OscdimgPath           = $null
  OscdimgLabel          = $null

  OutputIsoPath          = $null
  IsoBaseName            = $null

  StashedInstallWim      = $null

  WIM_VersionString      = $null
  WIM_BuildMajor         = $null
  WIM_ReleaseToken       = $null
  WIM_ArchCatalog        = $null  # x64/x86/arm64
  WIM_ArchToken          = $null  # amd64/x86/arm64
  WIM_OsFamily           = $null  # Windows 10/Windows 11
}

# Cancellation + child process tracking (for safe cleanup)
$script:Cancelled  = $false
$script:ChildProcs = New-Object System.Collections.Generic.List[System.Diagnostics.Process]

# ==============================
# Logging and failure helpers
# ==============================
function V([string]$Message) { Write-Verbose ("[{0}] {1}" -f (Get-Date -Format "s"), $Message) }

function Show-Usage {
  $name = if ($PSCommandPath) { Split-Path -Leaf $PSCommandPath } else { 'BundledWindowsIso.ps1' }
  Write-Host ""
  Write-Host "$name - Build <original>.bundled.iso from 1 ISO and optional MSU files in a folder." -ForegroundColor Cyan
  Write-Host ""
  Write-Host "USAGE:"
  Write-Host "  & '$name' [Folder] [-ISO <path>] [-DestISO <path>] [-UseSystemTemp] [-Verbose] [-DryRun]"
  Write-Host "  selection: -Home -Pro -Indices <spec>  (if none specified, uses all indexes)"
  Write-Host ""
  Write-Host "TOOL OVERRIDES:"
  Write-Host "  -UseADK               Prefer ADK tools (default if available)"
  Write-Host "  -UseSystem            Force system tools"
  Write-Host "  -dism <path>           Use explicit dism.exe"
  Write-Host "  -oscdimg <path>         Use explicit oscdimg.exe"
  Write-Host ""
}

function Fail([string]$Message, [int]$Code = 1) {
  Write-Host ""
  Write-Host "ERROR: $Message" -ForegroundColor Red
  Write-Verbose ("WorkRoot: {0}" -f $script:State.WorkRoot)
  Write-Verbose ("ISO: {0}" -f $script:State.IsoPath)
  if ($script:State.DismPath)    { Write-Verbose ("DISM: {0} ({1})" -f $script:State.DismPath, $script:State.DismLabel) }
  if ($script:State.OscdimgPath) { Write-Verbose ("OSCDIMG: {0} ({1})" -f $script:State.OscdimgPath, $script:State.OscdimgLabel) }
  Show-Usage
  exit $Code
}

function Assert-Admin {
  # DISM servicing + mounting requires elevation
  $id = [Security.Principal.WindowsIdentity]::GetCurrent()
  $p  = New-Object Security.Principal.WindowsPrincipal($id)
  if (-not $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Fail "Please run PowerShell as Administrator."
  }
}

function Assert-NotCancelled {
  if ($script:Cancelled) {
    throw [System.OperationCanceledException]::new("Cancelled by user (Ctrl+C).")
  }
}

# ==============================
# External process runner (tracked)
# ==============================
function Invoke-External {
  param(
    [Parameter(Mandatory=$true)][string]$FilePath,
    [Parameter(Mandatory=$true)][string[]]$ArgumentList
  )

  Assert-NotCancelled

  $p = Start-Process -FilePath $FilePath -ArgumentList $ArgumentList -PassThru -NoNewWindow
  $script:ChildProcs.Add($p) | Out-Null

  try {
    while (-not $p.HasExited) {
      Start-Sleep -Milliseconds 200
      try { $p.Refresh() } catch {}
    }
  } catch [System.Management.Automation.PipelineStoppedException] {
    $script:Cancelled = $true
    try { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue } catch {}
    throw
  }

  try { $p.Refresh() } catch {}
  try { return [int]$p.ExitCode } catch { return 0 }
}

function Stop-TrackedChildren {
  # Stop only the processes this script started (never try to kill wimserv.exe)
  foreach ($p in $script:ChildProcs) {
    try { $p.Refresh(); if (-not $p.HasExited) { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue } } catch {}
  }
}

# ==============================
# Tool resolution (ADK first, then system) with overrides
# ==============================
function Find-AdkToolPath {
  param(
    [Parameter(Mandatory=$true)][ValidateSet('dism','oscdimg')][string]$Tool
  )

  $root = $Config.AdkDeploymentToolsRoot
  if (-not (Test-Path $root)) { return $null }

  foreach ($arch in $Config.AdkToolArch) {
    $p = switch ($Tool) {
      'dism'    { Join-Path $root (Join-Path $arch 'DISM\dism.exe') }
      'oscdimg' { Join-Path $root (Join-Path $arch 'Oscdimg\oscdimg.exe') }
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

  if ($UseADK -and $UseSystem) { Fail "Use only one of -UseADK or -UseSystem." }

  # --- DISM selection ---
  if ($ExplicitDism) {
    if (-not (Test-Path $ExplicitDism)) { Fail "Specified -dism path not found: $ExplicitDism" }
    $script:State.DismPath  = $ExplicitDism
    $script:State.DismLabel = "Explicit"
  } elseif ($UseSystem) {
    $sys = "$env:windir\System32\dism.exe"
    if (-not (Test-Path $sys)) { Fail "System DISM not found at: $sys" }
    $script:State.DismPath  = $sys
    $script:State.DismLabel = "System(forced)"
  } else {
    $adk = Find-AdkToolPath -Tool dism
    if ($adk) {
      $script:State.DismPath  = $adk
      $script:State.DismLabel = "ADK(preferred)"
    } else {
      $sys = "$env:windir\System32\dism.exe"
      if (-not (Test-Path $sys)) { Fail "System DISM not found at: $sys" }
      $script:State.DismPath  = $sys
      $script:State.DismLabel = "System(fallback)"
    }
  }

  # --- OSCDIMG selection ---
  if ($ExplicitOscdimg) {
    if (-not (Test-Path $ExplicitOscdimg)) { Fail "Specified -oscdimg path not found: $ExplicitOscdimg" }
    $script:State.OscdimgPath  = $ExplicitOscdimg
    $script:State.OscdimgLabel = "Explicit"
  } elseif ($UseSystem) {
    $cmd = Get-Command oscdimg.exe -ErrorAction SilentlyContinue
    if (-not $cmd) { Fail "oscdimg.exe not found on PATH (install ADK Deployment Tools or use -oscdimg <path>)." }
    $script:State.OscdimgPath  = $cmd.Source
    $script:State.OscdimgLabel = "PATH(forced)"
  } else {
    $adk = Find-AdkToolPath -Tool oscdimg
    if ($adk) {
      $script:State.OscdimgPath  = $adk
      $script:State.OscdimgLabel = "ADK(preferred)"
    } else {
      $cmd = Get-Command oscdimg.exe -ErrorAction SilentlyContinue
      if (-not $cmd) { Fail "oscdimg.exe not found (install ADK Deployment Tools or use -oscdimg <path>)." }
      $script:State.OscdimgPath  = $cmd.Source
      $script:State.OscdimgLabel = "PATH(fallback)"
    }
  }

  Write-Host ("Using DISM:    {0} ({1})" -f $script:State.DismPath, $script:State.DismLabel) -ForegroundColor Cyan
  Write-Host ("Using OSCDIMG: {0} ({1})" -f $script:State.OscdimgPath, $script:State.OscdimgLabel) -ForegroundColor Cyan
}

# ==============================
# Work folder creation (no WinIsoBundler_ prefix)
# ==============================
function Get-UniqueWorkRoot([string]$WorkBase, [string]$IsoBase) {
  # We are already in _WinIsoBundlerWork, so WorkRoot should just be the ISO base name.
  $candidate = Join-Path $WorkBase $IsoBase
  if (-not (Test-Path $candidate)) { return $candidate }

  # If it already exists, pick a stable numbered suffix (so reruns don't stomp each other)
  for ($i = 2; $i -le 99; $i++) {
    $cand2 = Join-Path $WorkBase ("{0}_{1}" -f $IsoBase, $i)
    if (-not (Test-Path $cand2)) { return $cand2 }
  }

  # Worst-case fallback
  return Join-Path $WorkBase ("{0}_{1:yyyyMMdd_HHmmss}" -f $IsoBase, (Get-Date))
}

function Initialize-WorkPaths([string]$BaseFolder, [string]$IsoPath, [switch]$UseSystemTemp) {
  $base = if ($UseSystemTemp) { [System.IO.Path]::GetTempPath() } else { Join-Path $BaseFolder $Config.WorkParentSubfolder }
  New-Item -ItemType Directory -Path $base -Force | Out-Null

  $isoBase = [IO.Path]::GetFileNameWithoutExtension($IsoPath)
  $workRoot = Get-UniqueWorkRoot -WorkBase $base -IsoBase $isoBase

  $script:State.WorkBase    = $base
  $script:State.WorkRoot    = $workRoot
  $script:State.IsoRoot     = Join-Path $workRoot $Config.WorkIsoSubdir
  $script:State.InstallRoot = Join-Path $workRoot $Config.WorkInstallSubdir
  $script:State.MountRoot   = Join-Path $workRoot $Config.WorkMountSubdir
  $script:State.LogsRoot    = Join-Path $workRoot $Config.WorkLogsSubdir
  $script:State.ScratchRoot = Join-Path $workRoot $Config.WorkScratchSubdir
  $script:State.MetaPath    = Join-Path $workRoot $Config.MetaFileName
  $script:State.IsoBaseName = $isoBase
}

# ==============================
# Read-only mitigation (ONLY for install.wim / install.esd)
# ==============================
function Clear-ReadOnly([string]$Path) {
  if (-not (Test-Path $Path)) { return }
  V "Clearing Read-only: $Path"
  try { & attrib -R "$Path" | Out-Null } catch {}
}

function Clear-ReadOnly-InstallFiles([string]$IsoRoot) {
  # We only clear read-only on the WIM/ESD files DISM needs write access to.
  $wimCandidate = Join-Path $IsoRoot 'sources\install.wim'
  $esdCandidate = Join-Path $IsoRoot 'sources\install.esd'
  if (Test-Path $wimCandidate) { Clear-ReadOnly -Path $wimCandidate }
  if (Test-Path $esdCandidate) { Clear-ReadOnly -Path $esdCandidate }
}

# ==============================
# ISO operations (mount + copy)
# ==============================
function Mount-Iso([string]$IsoPath) {
  Assert-NotCancelled

  $img = Mount-DiskImage -ImagePath $IsoPath -PassThru
  $vol = $img | Get-Volume -ErrorAction SilentlyContinue

  if (-not $vol -or -not $vol.DriveLetter) {
    Dismount-DiskImage -ImagePath $IsoPath -ErrorAction SilentlyContinue | Out-Null
    $img = Mount-DiskImage -ImagePath $IsoPath -PassThru
    $vol = $img | Get-Volume
  }

  if (-not $vol -or -not $vol.DriveLetter) {
    Dismount-DiskImage -ImagePath $IsoPath -ErrorAction SilentlyContinue | Out-Null
    Fail "Failed to mount ISO or resolve a drive letter."
  }

  V "Mounted ISO drive: $($vol.DriveLetter):"
  return @{ Image=$img; Drive="$($vol.DriveLetter):" }
}

function Copy-IsoContents([string]$SrcDrive, [string]$DstFolder) {
  # Copy ISO contents to a writable working folder using robocopy.
  # Quiet unless -Verbose; quiet output goes to LOGS\robocopy.log.
  Assert-NotCancelled
  New-Item -ItemType Directory -Path $DstFolder -Force | Out-Null

  $isVerbose = ($VerbosePreference -eq 'Continue')

  New-Item -ItemType Directory -Path $script:State.LogsRoot -Force | Out-Null
  $logPath = Join-Path $script:State.LogsRoot "robocopy.log"

  $args = @("$SrcDrive\", "$DstFolder\", "*.*") + $Config.RobocopyArgsBase
  if (-not $isVerbose) { $args += $Config.RobocopyArgsQuiet }

  if (-not $isVerbose) {
    & robocopy @args *> $logPath
  } else {
    & robocopy @args
  }

  $rc = [int]$LASTEXITCODE
  if ($rc -ge 8) { Fail "Robocopy failed with exit code $rc. See log: $logPath" }
}

function Get-InstallImagePathFromMountedIso([string]$IsoDrive) {
  $sources = Join-Path $IsoDrive 'sources'
  $wim = Join-Path $sources 'install.wim'
  $esd = Join-Path $sources 'install.esd'
  if (Test-Path $wim) { return $wim }
  if (Test-Path $esd) { return $esd }
  return $null
}

# ==============================
# Image inspection (version/arch detection)
# ==============================
function Get-ImageInfoFromInstallFile([string]$InstallPath) {
  # Parse ONLY "Details for image" section to avoid capturing DISM header "Version:" line.
  $out = & $script:State.DismPath /Get-WimInfo /WimFile:"$InstallPath" /Index:1
  if ($LASTEXITCODE -ne 0) { return $null }

  $name=$null; $arch=$null; $ver=$null
  $inDetails=$false
  foreach ($line in $out) {
    if (-not $inDetails -and $line -match '^\s*Details\s+for\s+image\s*:') { $inDetails = $true; continue }
    if (-not $inDetails) { continue }

    if (-not $name -and $line -match '^\s*Name\s*:\s*(.+)\s*$') { $name = $Matches[1].Trim(); continue }
    if (-not $arch -and $line -match '^\s*Architecture\s*:\s*(.+)\s*$') { $arch = $Matches[1].Trim(); continue }
    if (-not $ver  -and $line -match '^\s*Version\s*:\s*(.+)\s*$') { $ver  = $Matches[1].Trim(); continue }

    if ($name -and $arch -and $ver) { break }
  }
  if (-not $name -or -not $arch -or -not $ver) { return $null }
  [pscustomobject]@{ Name=$name; Architecture=$arch; Version=$ver }
}

function Convert-ImageArchToCatalogArch([string]$Arch) {
  $a = ($Arch + '').ToLowerInvariant()
  if ($a -eq 'x64' -or $a -eq 'amd64') { return 'x64' }
  if ($a -eq 'x86') { return 'x86' }
  if ($a -eq 'arm64') { return 'arm64' }
  return $null
}

function Convert-ImageArchToUnattendArch([string]$Arch) {
  $a = ($Arch + '').ToLowerInvariant()
  if ($a -eq 'x64' -or $a -eq 'amd64') { return 'amd64' }
  if ($a -eq 'x86' -or $a -eq 'i386') { return 'x86' }
  if ($a -eq 'arm64') { return 'arm64' }
  return 'amd64'
}

function Guess-Win11ReleaseFromBuild([string]$VersionString) {
  $m = [regex]::Match($VersionString, '10\.0\.(\d+)')
  if (-not $m.Success) { return $Config.DefaultWin11Release }
  $build = [int]$m.Groups[1].Value
  if ($build -ge 26200) { return '25H2' }
  if ($build -ge 26100) { return '24H2' }
  if ($build -ge 22631) { return '23H2' }
  if ($build -ge 22621) { return '22H2' }
  return $Config.DefaultWin11Release
}

# ==============================
# Ensure install.wim exists in extracted ISO (convert ESD if needed)
# ==============================
function Ensure-InstallWimInExtractedIso([string]$IsoRoot) {
  $srcDir = Join-Path $IsoRoot 'sources'
  $wim = Join-Path $srcDir 'install.wim'
  $esd = Join-Path $srcDir 'install.esd'

  if (Test-Path $wim) { return $wim }
  if (-not (Test-Path $esd)) { Fail "Neither install.wim nor install.esd found in sources folder." }

  Write-Host "Converting ISO\sources\install.esd to ISO\sources\install.wim..." -ForegroundColor Cyan

  $info = & $script:State.DismPath /Get-ImageInfo /ImageFile:"$esd"
  if ($LASTEXITCODE -ne 0) { Fail "DISM failed to read install.esd." }

  $indexes = @()
  foreach ($line in $info) {
    if ($line -match '^\s*Index\s*:\s*(\d+)\s*$') { $indexes += [int]$Matches[1] }
  }
  if ($indexes.Count -lt 1) { Fail "Could not parse indexes from install.esd." }

  foreach ($idx in $indexes) {
    Assert-NotCancelled
    $rc = Invoke-External -FilePath $script:State.DismPath -ArgumentList @(
      "/Export-Image",
      "/SourceImageFile:$esd",
      "/SourceIndex:$idx",
      "/DestinationImageFile:$wim",
      "/Compress:max",
      "/CheckIntegrity"
    )
    if ($rc -ne 0) { Fail "DISM Export-Image failed for SourceIndex $idx (exit $rc)." }
  }

  Rename-Item -Path $esd -NewName 'install.esd.bak' -Force
  Clear-ReadOnly -Path $wim
  return $wim
}

function Get-WimPairs([string]$WimPath) {
  $info = & $script:State.DismPath /Get-WimInfo /WimFile:"$WimPath"
  if ($LASTEXITCODE -ne 0) { Fail "DISM Get-WimInfo failed for $WimPath." }
  $pairs = @()
  $cur = $null
  foreach ($line in $info) {
    if ($line -match '^\s*Index\s*:\s*(\d+)\s*$') { $cur = [int]$Matches[1]; continue }
    if ($cur -ne $null -and $line -match '^\s*Name\s*:\s*(.+)\s*$') {
      $pairs += [pscustomobject]@{ Index=$cur; Name=$Matches[1].Trim() }
      $cur = $null
    }
  }
  if ($pairs.Count -lt 1) { Fail "No indexes found in $WimPath." }
  return $pairs
}

# ==============================
# Injected scripts & autounattend.xml
# ==============================
function Write-PostInstallWingetScript([string]$DestPath) {
  # Create a small script that runs once at first logon (via autounattend FirstLogonCommands).
  $logName = $Config.WingetLogFile
  $wargs   = ($Config.WingetArgs -join ' ')
  $content = @"
# $($Config.PostInstallScriptName)
`$log = Join-Path `$env:ProgramData '$logName'
"[[`$(Get-Date -Format s)]] Starting winget upgrade --all" | Out-File -FilePath `$log -Append

try {
  if (-not (Get-Command winget.exe -ErrorAction SilentlyContinue)) {
    "[[`$(Get-Date -Format s)]] winget.exe not found; skipping." | Out-File -FilePath `$log -Append
    return
  }
  winget upgrade $wargs
  "[[`$(Get-Date -Format s)]] Completed winget upgrade --all" | Out-File -FilePath `$log -Append
}
catch {
  "[[`$(Get-Date -Format s)]] ERROR: `$(`$_.Exception.Message)" | Out-File -FilePath `$log -Append
}
"@
  Set-Content -Path $DestPath -Value $content -Encoding ASCII
}

function Write-Autounattend([string]$IsoRoot, [string]$ArchToken) {
  # Write autounattend.xml to ISO root to run the post-install script at first logon.
  $path = Join-Path $IsoRoot $Config.AutounattendName
  if (Test-Path $path) { Copy-Item $path "$path.bak" -Force }

  $ps = "%WINDIR%\Setup\Scripts\$($Config.PostInstallScriptName)"
  $driversPath = $Config.DriverFolderUnattendPath

  $xml = @"
<?xml version="1.0" encoding="utf-8"?>
<unattend xmlns="urn:schemas-microsoft-com:unattend">
  <settings pass="windowsPE">
    <component name="Microsoft-Windows-PnpCustomizationsWinPE"
               processorArchitecture="$ArchToken"
               publicKeyToken="31bf3856ad364e35"
               language="neutral"
               versionScope="nonSxS"
               xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State"
               xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
      <DriverPaths>
        <PathAndCredentials wcm:action="add" wcm:keyValue="1">
          <Path>$driversPath</Path>
        </PathAndCredentials>
      </DriverPaths>
    </component>
  </settings>
  <settings pass="oobeSystem">
    <component name="Microsoft-Windows-Shell-Setup"
               processorArchitecture="$ArchToken"
               publicKeyToken="31bf3856ad364e35"
               language="neutral"
               versionScope="nonSxS"
               xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State"
               xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
      <FirstLogonCommands>
        <SynchronousCommand wcm:action="add">
          <Order>1</Order>
          <Description>Run WinGet upgrade --all</Description>
          <CommandLine>powershell.exe -ExecutionPolicy Bypass -File $ps</CommandLine>
        </SynchronousCommand>
      </FirstLogonCommands>
    </component>
  </settings>
</unattend>
"@
  Set-Content -Path $path -Value $xml -Encoding ASCII
}

# ==============================
# MSCatalogLTS download (no MSU validation; download then apply)
# ==============================
function Ensure-MSCatalogLTS {
  Ensure-Tls12
  if (-not (Get-Module -ListAvailable -Name MSCatalogLTS)) {
    V "Installing MSCatalogLTS module (CurrentUser)..."
    Install-Module -Name MSCatalogLTS -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
  }
  Import-Module MSCatalogLTS -ErrorAction Stop
}

function Select-LatestCatalogUpdate([object[]]$Results, [string[]]$IncludeTokens, [string[]]$ExcludeTokens) {
  $filtered = @()
  foreach ($r in $Results) {
    $t = ($r.Title + '')
    if (-not $t) { continue }
    $tl = $t.ToLowerInvariant()

    if ($tl -notmatch 'cumulative update') { continue }

    # drop previews/.net/etc
    foreach ($tok in $Config.ExcludeTitleTokens) {
      if ($tl -like ("*" + $tok.ToLowerInvariant() + "*")) { $t = $null; break }
    }
    if (-not $t) { continue }

    # drop server items
    if ($ExcludeTokens) {
      foreach ($tok in $ExcludeTokens) {
        if ($tl -like ("*" + $tok.ToLowerInvariant() + "*")) { $t = $null; break }
      }
      if (-not $t) { continue }
    }

    # require includes
    if ($IncludeTokens) {
      $ok = $true
      foreach ($tok in $IncludeTokens) {
        if ($tl -notlike ("*" + $tok.ToLowerInvariant() + "*")) { $ok = $false; break }
      }
      if (-not $ok) { continue }
    }

    $filtered += $r
  }

  if ($filtered.Count -gt 0) { return $filtered[0] }
  return $null
}

function Download-LatestCumulativeUpdateMsu([string]$Folder, [string[]]$SearchTexts, [string]$CatalogArch, [string[]]$IncludeTokens) {
  Ensure-MSCatalogLTS

  $results = $null
  foreach ($q in $SearchTexts) {
    Write-Host "Searching Microsoft Update Catalog (MSCatalogLTS)..." -ForegroundColor Cyan
    Write-Host "  Search: $q" -ForegroundColor Cyan
    Write-Host "  Arch:   $CatalogArch" -ForegroundColor Cyan
    $results = Get-MSCatalogUpdate -Search $q -Architecture $CatalogArch
    if ($results) { break }
  }
  if (-not $results) { Fail "No updates found for search strings: $($SearchTexts -join ' | ')" }

  $u = Select-LatestCatalogUpdate -Results $results -IncludeTokens $IncludeTokens -ExcludeTokens $Config.ExcludeServerTitleTokens
  if (-not $u) { Fail "Could not select an update from catalog results (after filtering)." }

  Write-Host "Selected update title: $($u.Title)" -ForegroundColor Cyan
  Save-MSCatalogUpdate -Update $u -Destination $Folder | Out-Null
}

function Ensure-MsusPresentOrDownload([string]$Folder, [string]$OsFamily, [string]$ReleaseToken, [string]$CatalogArch, [bool]$IsCheckpointCapable) {
  $msuFiles = @(Get-ChildItem -LiteralPath $Folder -Filter "*.msu" -File -ErrorAction SilentlyContinue)
  if ($msuFiles.Count -gt 0) { return $true }

  if ($OsFamily -eq 'Windows 10') {
    $searches = @(
      "Cumulative Update for Windows 10 Version $ReleaseToken for $CatalogArch-based Systems",
      "Cumulative Update for Windows 10, version $ReleaseToken for $CatalogArch-based Systems"
    )
    $include = @("windows 10", "$CatalogArch-based systems", $ReleaseToken)
    Download-LatestCumulativeUpdateMsu -Folder $Folder -SearchTexts $searches -CatalogArch $CatalogArch -IncludeTokens $include
  } else {
    $searches = @(
      "Cumulative Update for Windows 11, version $ReleaseToken for $CatalogArch-based Systems",
      "Cumulative Update for Windows 11 Version $ReleaseToken for $CatalogArch-based Systems"
    )
    $include = @("windows 11", "$CatalogArch-based systems", $ReleaseToken)
    Download-LatestCumulativeUpdateMsu -Folder $Folder -SearchTexts $searches -CatalogArch $CatalogArch -IncludeTokens $include
  }

  # Best-effort checkpoint (KB5043080) for 24H2+
  if ($IsCheckpointCapable -and $OsFamily -eq 'Windows 11') {
    foreach ($kb in $Config.KnownCheckpointKBs24H2Plus) {
      $already = @(Get-ChildItem -LiteralPath $Folder -Filter "*$kb*.msu" -File -ErrorAction SilentlyContinue)
      if ($already.Count -gt 0) { continue }

      Write-Host "Attempting checkpoint prerequisite (best-effort): $kb" -ForegroundColor Cyan
      # Force Windows 11 Version 24H2 + arch; exclude server titles
      $include = @("windows 11", "version 24h2", "$CatalogArch-based systems", $kb.ToLowerInvariant())
      Download-LatestCumulativeUpdateMsu -Folder $Folder -SearchTexts @($kb) -CatalogArch $CatalogArch -IncludeTokens $include
    }
  }

  $msuFiles = @(Get-ChildItem -LiteralPath $Folder -Filter "*.msu" -File -ErrorAction SilentlyContinue)
  return ($msuFiles.Count -gt 0)
}

# ==============================
# Service WIM index (mount/add-package/commit)
# ==============================
function Service-WimIndex {
  param(
    [Parameter(Mandatory=$true)][string]$WimPath,
    [Parameter(Mandatory=$true)][int]$Index,
    [Parameter(Mandatory=$true)][System.IO.FileInfo[]]$Packages,
    [Parameter(Mandatory=$true)][string]$MountRoot,
    [Parameter(Mandatory=$true)][string]$PostInstallScriptSource,
    [Parameter(Mandatory=$true)][string]$LogsRoot,
    [Parameter(Mandatory=$true)][string]$ScratchRoot
  )

  Assert-NotCancelled

  # Use a per-index mount directory
  $mountDir = Join-Path $MountRoot ("idx_{0}" -f $Index)
  $mountLog = Join-Path $LogsRoot ("dism_mount_idx{0}.log" -f $Index)

  if (Test-Path $mountDir) { Remove-Item $mountDir -Recurse -Force -ErrorAction SilentlyContinue }
  New-Item -ItemType Directory -Path $mountDir -Force | Out-Null

  Write-Host "Mounting WIM index $Index..." -ForegroundColor Cyan

  $rc = Invoke-External -FilePath $script:State.DismPath -ArgumentList @(
    "/Mount-Image",
    "/ImageFile:$WimPath",
    "/Index:$Index",
    "/MountDir:$mountDir",
    "/ScratchDir:$ScratchRoot",
    "/LogPath:$mountLog"
  )
  if ($rc -ne 0) { Fail "Failed to mount WIM index $Index (exit $rc). See log: $mountLog" }

  try {
    # Inject first-logon script
    $setupScripts = Join-Path $mountDir 'Windows\Setup\Scripts'
    New-Item -ItemType Directory -Path $setupScripts -Force | Out-Null
    Copy-Item -Path $PostInstallScriptSource -Destination (Join-Path $setupScripts $Config.PostInstallScriptName) -Force

    # Add each MSU (no validation; let DISM handle)
    foreach ($pkg in $Packages) {
      Assert-NotCancelled
      Write-Host "  Adding package: $($pkg.Name)" -ForegroundColor DarkCyan

      $safeName = ($pkg.BaseName -replace '[^A-Za-z0-9._-]','_')
      $pkgLog  = Join-Path $LogsRoot ("dism_addpackage_idx{0}_{1}.log" -f $Index, $safeName)

      $rc2 = Invoke-External -FilePath $script:State.DismPath -ArgumentList @(
        "/Image:$mountDir",
        "/Add-Package",
        "/PackagePath:$($pkg.FullName)",
        "/ScratchDir:$ScratchRoot",
        "/LogPath:$pkgLog"
      )
      if ($rc2 -ne 0) { Fail "DISM Add-Package failed for $($pkg.Name) on index $Index (exit $rc2). See log: $pkgLog" }
    }
  }
  finally {
    Write-Host "Committing and unmounting index $Index..." -ForegroundColor Cyan
    $rc3 = Invoke-External -FilePath $script:State.DismPath -ArgumentList @(
      "/Unmount-Image",
      "/MountDir:$mountDir",
      "/Commit"
    )
    if ($rc3 -ne 0) { Fail "Failed to unmount/commit index $Index (exit $rc3)." }
  }
}

# ==============================
# Build ISO (oscdimg)
# ==============================
function Build-Iso([string]$IsoRoot, [string]$OutputIso) {
  $etfs = Join-Path $IsoRoot $Config.BootFileBIOS
  $efis = Join-Path $IsoRoot $Config.BootFileUEFI
  if (-not (Test-Path $etfs)) { Fail "Missing BIOS boot file: $etfs" }
  if (-not (Test-Path $efis)) { Fail "Missing UEFI boot file: $efis" }

  # BIOS+UEFI bootdata syntax
  $bootdata = "2#p0,e,b$etfs#pEF,e,b$efis"
  $args = @() + $Config.OscdimgFsArgs + @("-l$($Config.IsoVolumeLabel)", "-bootdata:$bootdata", $IsoRoot, $OutputIso)

  Write-Host "Building ISO: $OutputIso" -ForegroundColor Green
  $rc = Invoke-External -FilePath $script:State.OscdimgPath -ArgumentList $args
  if ($rc -ne 0) { Fail "oscdimg failed with exit code $rc" }
}

# ==============================
# DISM.EXE-only lock release + cleanup
# ==============================
function Release-WimLocks {
  # Use chosen DISM. Cleanup mount points and discard mounts under our mount root.
  $dism = $script:State.DismPath
  if (-not $dism -or -not (Test-Path $dism)) { $dism = "$env:windir\System32\dism.exe" }
  if (-not $script:State.MountRoot) { return }

  $out = & $dism /Get-MountedImageInfo 2>$null
  $mountDirs = @()
  if ($LASTEXITCODE -eq 0 -and $out) {
    foreach ($line in $out) {
      if ($line -match '^\s*Mount Dir\s*:\s*(.+)\s*$') {
        $mountDirs += $Matches[1].Trim()
      }
    }
  }

  foreach ($md in $mountDirs) {
    try {
      if ($md -and $md.ToLowerInvariant().StartsWith($script:State.MountRoot.ToLowerInvariant())) {
        & $dism /Unmount-Image /MountDir:"$md" /Discard | Out-Null
      }
    } catch {}
  }

  try { & $dism /Cleanup-MountPoints | Out-Null } catch {}
}

function Cleanup-Hardened {
  Stop-TrackedChildren
  Release-WimLocks
  try {
    if ($script:State.IsoWasMounted -and $script:State.IsoPath) {
      Dismount-DiskImage -ImagePath $script:State.IsoPath -ErrorAction SilentlyContinue | Out-Null
    }
  } catch {}
}

# ==============================
# Argument parsing
# ==============================
$FolderArg = $null
$UseSystemTemp = $false
$ShowHelp = $false
$VerboseSwitch = $false

$UseADK = $false
$UseSystem = $false
$ExplicitDism = $null
$ExplicitOscdimg = $null

$CleanWork = $false
$ShowIndices = $false
$SelectHome = $false
$SelectPro = $false
$IndicesSpec = $null

$IsoArg = $null
$DestIsoArg = $null

for ($i = 0; $i -lt $args.Count; $i++) {
  $a = $args[$i]
  switch -Regex ($a) {
    '^(?:-help|-h|-\?|/h|/\?)$' { $ShowHelp = $true; continue }
    '^(?:-UseSystemTemp)$'     { $UseSystemTemp = $true; continue }
    '^(?:-Verbose|-v)$'        { $VerboseSwitch = $true; continue }
    '^(?:-UseADK)$'            { $UseADK = $true; continue }
    '^(?:-UseSystem)$'         { $UseSystem = $true; continue }
    '^(?:-dism)$'              { $ExplicitDism = [string]$args[++$i]; continue }
    '^(?:-oscdimg)$'            { $ExplicitOscdimg = [string]$args[++$i]; continue }

    '^(?:-CleanWork)$'         { $CleanWork = $true; continue }
    '^(?:-ShowIndices)$'       { $ShowIndices = $true; continue }
    '^(?:-Show)$'              { $ShowIndices = $true; continue }

    '^(?:-Home)$'              { $SelectHome = $true; continue }
    '^(?:-Pro)$'               { $SelectPro = $true; continue }

    '^(?:-Indices)$' {
      if ($i + 1 -ge $args.Count) { Fail "-Indices requires a spec string" }
      $IndicesSpec = [string]$args[++$i]
      continue
    }

    '^(?:-ISO|-SrcISO)$' {
      if ($i + 1 -ge $args.Count) { Fail "-ISO requires a path" }
      $IsoArg = [string]$args[++$i]
      continue
    }
    '^(?:-DestISO)$' {
      if ($i + 1 -ge $args.Count) { Fail "-DestISO requires a path" }
      $DestIsoArg = [string]$args[++$i]
      continue
    }

    '^-{1,2}.*' { Fail "Unknown switch: $a" }
    default {
      if ($FolderArg) { Fail "Only one folder argument allowed. Extra value: $a" }
      $FolderArg = $a
    }
  }
}

if ($ShowHelp) { Show-Usage; exit 0 }
if (-not $FolderArg) { $FolderArg = (Get-Location).Path }
if ($VerboseSwitch) { $VerbosePreference = 'Continue' }

# ==============================
# Main
# ==============================
Assert-Admin
Resolve-Tools -ExplicitDism $ExplicitDism -ExplicitOscdimg $ExplicitOscdimg -UseADK:($UseADK) -UseSystem:($UseSystem)

try {
  $folderPath = (Resolve-Path -LiteralPath $FolderArg -ErrorAction SilentlyContinue).Path
  if (-not $folderPath) { Fail "Folder not found: $FolderArg" }

  # Resolve ISO
  if ($IsoArg) {
    $resolvedIso = (Resolve-Path -LiteralPath $IsoArg -ErrorAction SilentlyContinue).Path
    if (-not $resolvedIso) { Fail "ISO not found: $IsoArg" }
    $script:State.IsoPath = $resolvedIso
  } else {
    # Exactly one ISO in folder
    $isoFiles = Get-ChildItem -LiteralPath $folderPath -Filter "*.iso" -File | Where-Object { $_.Name -notlike "*$($Config.OutputIsoSuffix)" }
    if ($isoFiles.Count -ne 1) { Fail "Expected exactly 1 ISO in '$folderPath'. Found $($isoFiles.Count)." }
    $script:State.IsoPath = $isoFiles[0].FullName
  }

  # Init work paths
  Initialize-WorkPaths -BaseFolder $folderPath -IsoPath $script:State.IsoPath -UseSystemTemp:($UseSystemTemp)

  New-Item -ItemType Directory -Path $script:State.WorkRoot, $script:State.IsoRoot, $script:State.InstallRoot, $script:State.MountRoot, $script:State.LogsRoot, $script:State.ScratchRoot -Force | Out-Null
  Write-Host ("Work root: {0}" -f $script:State.WorkRoot) -ForegroundColor DarkGray

  # Determine if ISO was already extracted.
  # We treat it as "extracted" if setup.exe and sources folder exist.
  $setupExe = Join-Path $script:State.IsoRoot 'setup.exe'
  $sourcesDir = Join-Path $script:State.IsoRoot 'sources'
  $alreadyExtracted = (Test-Path $setupExe) -and (Test-Path $sourcesDir)

  if (-not $alreadyExtracted) {
    # Copy ISO contents only when needed (old behavior)
    $isoMount = Mount-Iso $script:State.IsoPath
    $script:State.IsoWasMounted = $true

    Write-Host "Copying ISO contents..." -ForegroundColor Cyan
    Copy-IsoContents -SrcDrive $isoMount.Drive -DstFolder $script:State.IsoRoot

    # Narrow read-only fix: only install.wim/esd
    Clear-ReadOnly-InstallFiles -IsoRoot $script:State.IsoRoot
  } else {
    Write-Host "ISO contents already extracted; skipping robocopy." -ForegroundColor DarkGray
  }

  # Ensure Drivers folder exists so it ends up in ISO
  $driversFolderOnIso = Join-Path $script:State.IsoRoot $Config.DriverFolderName
  New-Item -ItemType Directory -Path $driversFolderOnIso -Force | Out-Null

  # Mount ISO only for metadata when needed (old behavior)
  # If we already extracted, we can read install.* from extracted ISO.
  $installPath = $null
  $installPathExtractedWim = Join-Path $script:State.IsoRoot 'sources\install.wim'
  $installPathExtractedEsd = Join-Path $script:State.IsoRoot 'sources\install.esd'
  if (Test-Path $installPathExtractedWim) { $installPath = $installPathExtractedWim }
  elseif (Test-Path $installPathExtractedEsd) { $installPath = $installPathExtractedEsd }
  else { Fail "Could not find sources\install.wim or sources\install.esd in extracted ISO tree." }

  # Pull OS / build info from index 1 details
  Write-Host "Analyzing installation image metadata (install.wim)..." -ForegroundColor Cyan
  $imageInfo = Get-ImageInfoFromInstallFile -InstallPath $installPath
  if (-not $imageInfo) { Fail "Could not read image info from install.wim/esd." }

  $catalogArch = Convert-ImageArchToCatalogArch -Arch $imageInfo.Architecture
  if (-not $catalogArch) { Fail "Unsupported ISO architecture: $($imageInfo.Architecture)" }
  $unattendArch = Convert-ImageArchToUnattendArch -Arch $imageInfo.Architecture
  $script:State.WIM_ArchCatalog = $catalogArch
  $script:State.WIM_ArchToken = $unattendArch
  $script:State.WIM_VersionString = $imageInfo.Version

  $nameLower = ($imageInfo.Name + '').ToLowerInvariant()
  if ($nameLower -match 'windows 10') { $script:State.WIM_OsFamily = 'Windows 10' }
  elseif ($nameLower -match 'windows 11') { $script:State.WIM_OsFamily = 'Windows 11' }
  else { $script:State.WIM_OsFamily = 'Windows 11' }

  if ($script:State.WIM_OsFamily -eq 'Windows 11') {
    $script:State.WIM_ReleaseToken = Guess-Win11ReleaseFromBuild -VersionString $script:State.WIM_VersionString
  } else {
    $script:State.WIM_ReleaseToken = $Config.DefaultWin10Release
  }

  # Build number for checkpoint capability
  $m = [regex]::Match($script:State.WIM_VersionString, '10\.0\.(\d+)')
  if ($m.Success) { $script:State.WIM_BuildMajor = [int]$m.Groups[1].Value }

  Write-Host "ISO OS detected: $($script:State.WIM_OsFamily)" -ForegroundColor Cyan
  Write-Host "ISO architecture detected: $($imageInfo.Architecture) (catalog: $catalogArch)" -ForegroundColor Cyan
  if ($script:State.WIM_OsFamily -eq 'Windows 11') { Write-Host "Windows 11 release line (best effort): $($script:State.WIM_ReleaseToken)" -ForegroundColor Cyan }

  # Show indexes if requested
  $pairs = Get-WimPairs -WimPath (Ensure-InstallWimInExtractedIso -IsoRoot $script:State.IsoRoot)
  if ($ShowIndices) {
    $pairs | ForEach-Object { Write-Host ("{0,2} {1}" -f $_.Index, $_.Name) }
    return
  }

  # Pick indices: Home/Pro/Indices, else all
  $maxIndex = ($pairs | Measure-Object -Property Index -Maximum).Maximum
  $nameMap = @{}
  foreach ($p in $pairs) { $nameMap[[int]$p.Index] = $p.Name }

  $selected = @()
  if ($SelectHome) { $selected += (Resolve-LabelTokenToIndices -Pairs $pairs -Token "Home") }
  if ($SelectPro)  { $selected += (Resolve-LabelTokenToIndices -Pairs $pairs -Token "Pro") }

  if ($IndicesSpec -and $IndicesSpec.Trim() -ne "") {
    $parsed = Parse-IndicesSpec -Spec $IndicesSpec -MaxIndex $maxIndex
    if ($parsed.Invalid.Count -gt 0) { Fail ("Invalid -Indices tokens: {0}" -f ($parsed.Invalid -join ", ")) }
    $selected += $parsed.Selected
  }

  if ($selected.Count -lt 1) { $selected = 1..$maxIndex } else { $selected = @($selected | Sort-Object -Unique) }

  Write-Host "Selected source indices (index : name):" -ForegroundColor Cyan
  foreach ($idx in $selected) {
    $nm = $nameMap[[int]$idx]; if (-not $nm) { $nm = "<unknown>" }
    Write-Host ("  {0}: {1}" -f $idx, $nm) -ForegroundColor Cyan
  }

  # Output ISO path
  if ($DestIsoArg) { $outIso = $DestIsoArg }
  else {
    $baseName = [IO.Path]::GetFileNameWithoutExtension($script:State.IsoPath)
    $outIso = Join-Path $folderPath ($baseName + $Config.OutputIsoSuffix)
  }

  # Discover local MSUs (if present)
  $msuFiles = @(Get-ChildItem -LiteralPath $folderPath -Filter "*.msu" -File -ErrorAction SilentlyContinue)
  V ("MSU files found: {0}" -f $msuFiles.Count)

  # Download MSU if none provided (old behavior)
  if ($msuFiles.Count -lt 1) {
    $isCheckpointCapable = ($script:State.WIM_BuildMajor -ge 26100)
    $ok = Ensure-MsusPresentOrDownload -Folder $folderPath -OsFamily $script:State.WIM_OsFamily -ReleaseToken $script:State.WIM_ReleaseToken -CatalogArch $catalogArch -IsCheckpointCapable:$isCheckpointCapable
    if (-not $ok) { Fail "MSU download did not produce any .msu files in '$folderPath'." }
    $msuFiles = @(Get-ChildItem -LiteralPath $folderPath -Filter "*.msu" -File -ErrorAction SilentlyContinue)
  }

  # Apply MSUs in simple name order (old behavior)
  $sortedMsus = @($msuFiles | Sort-Object Name)
  Write-Host ("MSUs available for servicing: {0}" -f $sortedMsus.Count) -ForegroundColor Cyan

  # Generate PostInstallWinget.ps1 (temporary file)
  $postInstallTemp = Join-Path $script:State.MountRoot $Config.PostInstallScriptName
  Write-PostInstallWingetScript -DestPath $postInstallTemp

  # Ensure install.wim exists in extracted ISO (convert ESD if needed)
  $wimPath = Ensure-InstallWimInExtractedIso -IsoRoot $script:State.IsoRoot
  Clear-ReadOnly -Path $wimPath

  # Place autounattend.xml in ISO root
  Write-Autounattend -IsoRoot $script:State.IsoRoot -ArchToken $unattendArch

  # Service each selected index (old behavior)
  foreach ($idx in $selected) {
    Service-WimIndex -WimPath $wimPath -Index $idx -Packages $sortedMsus -MountRoot $script:State.MountRoot -PostInstallScriptSource $postInstallTemp -LogsRoot $script:State.LogsRoot -ScratchRoot $script:State.ScratchRoot
  }

  # Build output ISO
  Build-Iso -IsoRoot $script:State.IsoRoot -OutputIso $outIso

  Write-Host "SUCCESS: Created $outIso" -ForegroundColor Green
  Write-Host "WinGet runs at first logon. Log: %ProgramData%\$($Config.WingetLogFile)" -ForegroundColor Green
}
catch [System.Management.Automation.PipelineStoppedException] {
  $script:Cancelled = $true
  Write-Host "Ctrl+C detected. Cancelling and cleaning up..." -ForegroundColor Yellow
}
catch [System.OperationCanceledException] {
  $script:Cancelled = $true
  Write-Host "Cancelled. Cleaning up..." -ForegroundColor Yellow
}
finally {
  Cleanup-Hardened
}
