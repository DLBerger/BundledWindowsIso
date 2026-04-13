<#
.SYNOPSIS
Creates a bundled Windows installation ISO by extracting an input ISO, optionally rebuilding install.wim from selected editions, servicing the image(s), and generating a new bootable ISO.

.DESCRIPTION
This script operates on a folder that contains a Windows ISO (excluding *.bundled.iso). It extracts the ISO into a stable work directory, prepares an installation image, optionally services it, writes an autounattend.xml, and builds a new ISO.

Help display:
- Use -h, -help, -? (or /?) to display this embedded help.
- The script displays help by calling Get-Help -Path <this script> -Full. Comment-based help is supported for scripts and functions. [1](https://stackoverflow.com/questions/37889252/pass-string-variable-to-function-argument-as-comma-separated-list)[2](https://catalog.update.microsoft.com/Search.aspx?q=25h2)
- Use -Full to display all sections reliably, including NOTES. [3](https://stackoverflow.com/questions/4988226/how-do-i-pass-multiple-parameters-into-a-function-in-powershell)

Drivers folder:
- The script creates a folder named "Drivers" at the root of the extracted ISO work tree (same level as setup.exe) so it exists in the final ISO.
- autounattend.xml includes a windowsPE driver search path using Microsoft-Windows-PnpCustomizationsWinPE -> DriverPaths -> PathAndCredentials.
  The path used is: %configsetroot%\Drivers
  DriverPaths/PathAndCredentials is the standard container/list structure for specifying driver search paths in an answer file. 

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
  - The script does not service the image(s)
  - The script does not merge/apply MSU packages
- To force a rebuild (and subsequent servicing), explicitly specify indices (for example: -Pro or -Indices 6,8,10).

DryRun behavior:
- With -DryRun, the script completes PREP actions needed to stage the work tree and then prints what would happen for post-PREP actions.

.PARAMETER Folder
Optional. Folder to process. If omitted, the current directory is used.

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

.EXAMPLE
PS> .\BundledWindowsIso.ps1 -UpdateISO
Reuses the existing work folder and performs no rebuild/servicing/MSU actions (no indices were specified).

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
- For help display, the script should resolve its own file path once (for example, $ScriptPath) and reuse it when calling Get-Help -Path $ScriptPath -Full. [1](https://stackoverflow.com/questions/37889252/pass-string-variable-to-function-argument-as-comma-separated-list)[2](https://catalog.update.microsoft.com/Search.aspx?q=25h2)
- The Drivers folder is intended for INF-based drivers (subfolders allowed) and is referenced via the answer file driver path. 

#>

# ==============================
# Script identity (resolved once)
# ==============================
$script:ScriptPath = $PSCommandPath
if (-not $script:ScriptPath) { $script:ScriptPath = $MyInvocation.MyCommand.Path }

# ==============================
# Configuration
# ==============================
$Config = [ordered]@{
  OutputIsoSuffix        = '.bundled.iso'

  WorkParentSubfolder    = '_WinIsoBundlerWork'
  WorkPrefix             = 'WinIsoBundler_'
  WorkIsoSubdir          = 'ISO'
  WorkInstallSubdir      = 'INSTALL'
  WorkMountSubdir        = 'MOUNT'
  WorkLogsSubdir         = 'LOGS'
  WorkScratchSubdir      = 'SCRATCH'
  MetaFileName           = 'bundler.meta.txt'

  DriverFolderName       = 'Drivers'
  DriverFolderUnattendPath = $null

  RobocopyArgsBase       = @('/E','/R:2','/W:2','/DCOPY:DA','/COPY:DAT')
  RobocopyArgsQuiet      = @('/NFL','/NDL','/NJH','/NJS','/NC','/NS','/NP')

  BootFileBIOS           = 'boot\etfsboot.com'
  BootFileUEFI           = 'efi\microsoft\boot\efisys.bin'
  IsoVolumeLabel         = 'WIN_BUNDLED'
  OscdimgFsArgs          = @('-m','-o','-u2','-udfver102')

  AutounattendName       = 'autounattend.xml'
  PostInstallScriptName  = 'PostInstallWinget.ps1'
  WingetLogFile          = 'WingetPostInstall.log'
  WingetArgs             = @('--all','--include-unknown','--silent','--accept-package-agreements','--accept-source-agreements')

  ExcludeTitleTokens     = @('preview', '.net', 'dynamic', 'safe os', 'setup dynamic')
  MaxCatalogQueries      = 8

  AdkToolArch            = @('amd64','arm64')
  AdkDeploymentToolsRoot = "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools"
}
$Config.DriverFolderUnattendPath = "%configsetroot%\$($Config.DriverFolderName)"

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
  MetaPath          = $null

  DismPath          = $null
  DismLabel         = $null
  OscdimgPath       = $null
  OscdimgLabel      = $null

  OutputIsoPath     = $null
  IsoBaseName       = $null

  StashedInstallWim = $null

  WIM_VersionString = $null
  WIM_BuildMajor    = $null
  WIM_ReleaseToken  = $null
  WIM_ArchCatalog   = $null
}

$script:Cancelled   = $false
$script:DryRun      = $false
$script:DryRunPhase = "afterprep" # "prep" or "afterprep"
$script:ChildProcs  = New-Object System.Collections.Generic.List[System.Diagnostics.Process]

# ==============================
# Helpers
# ==============================
function V([string]$Message) { Write-Verbose ("[{0}] {1}" -f (Get-Date -Format "s"), $Message) }

function Show-Usage {
  $name = if ($script:ScriptPath) { Split-Path -Leaf $script:ScriptPath } else { 'BundledWindowsIso.ps1' }
  Write-Host ""
  Write-Host "$name - use embedded help:" -ForegroundColor Cyan
  Write-Host "  & '$name' -h" -ForegroundColor Cyan
  Write-Host "  Get-Help -Path '$name' -Full" -ForegroundColor Cyan
  Write-Host ""
}

function Show-HelpAndExit {
  $p = $script:ScriptPath
  if (-not $p) { $p = $MyInvocation.MyCommand.Path }
  try {
    if ($p -and (Test-Path -LiteralPath $p)) {
      Get-Help -Path (Resolve-Path -LiteralPath $p).Path -Full
      exit 0
    }
  } catch {}
  Show-Usage
  exit 0
}

function Fail([string]$Message, [int]$Code = 1) {
  Write-Host ""
  Write-Host "ERROR: $Message" -ForegroundColor Red
  Show-Usage
  exit $Code
}

function Assert-Admin {
  $id = [Security.Principal.WindowsIdentity]::GetCurrent()
  $p  = New-Object Security.Principal.WindowsPrincipal($id)
  if (-not $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Fail "Please run PowerShell as Administrator."
  }
}

function Assert-NotCancelled {
  if ($script:Cancelled) { throw [System.OperationCanceledException]::new("Cancelled by user (Ctrl+C).") }
}

function In-AfterPrepDryRun { return ($script:DryRun -and $script:DryRunPhase -eq 'afterprep') }

function Invoke-Step([string]$What, [scriptblock]$Action) {
  if (In-AfterPrepDryRun) {
    Write-Host ("[DryRun] {0}" -f $What) -ForegroundColor Yellow
    return $null
  }
  V $What
  return & $Action
}

# ==============================
# External process runner
# ==============================
function Invoke-External {
  param(
    [Parameter(Mandatory=$true)][string]$FilePath,
    [Parameter(Mandatory=$true)][string[]]$ArgumentList
  )

  Assert-NotCancelled

  if (In-AfterPrepDryRun) {
    Write-Host ("[DryRun] EXEC: {0} {1}" -f $FilePath, ($ArgumentList -join ' ')) -ForegroundColor Yellow
    return 0
  }

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
  foreach ($p in $script:ChildProcs) {
    try { $p.Refresh(); if (-not $p.HasExited) { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue } } catch {}
  }
}

function Invoke-DismRead {
  param([Parameter(Mandatory=$true)][string[]]$Args)
  $out = & $script:State.DismPath @Args
  return ,$out
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

  if ($UseADK -and $UseSystem) { Fail "Use only one of -UseADK or -UseSystem." }

  if ($ExplicitDism) {
    if (-not (Test-Path $ExplicitDism)) { Fail "Specified -dism path not found: $ExplicitDism" }
    $script:State.DismPath = $ExplicitDism; $script:State.DismLabel = "Explicit"
  } else {
    $adk = if (-not $UseSystem) { Find-AdkToolPath dism } else { $null }
    if ($adk) { $script:State.DismPath = $adk; $script:State.DismLabel = "ADK" }
    else { $script:State.DismPath = "$env:windir\System32\dism.exe"; $script:State.DismLabel = "System" }
    if (-not (Test-Path $script:State.DismPath)) { Fail "DISM not found." }
  }

  if ($ExplicitOscdimg) {
    if (-not (Test-Path $ExplicitOscdimg)) { Fail "Specified -oscdimg path not found: $ExplicitOscdimg" }
    $script:State.OscdimgPath = $ExplicitOscdimg; $script:State.OscdimgLabel = "Explicit"
  } else {
    $adk = if (-not $UseSystem) { Find-AdkToolPath oscdimg } else { $null }
    if ($adk) { $script:State.OscdimgPath = $adk; $script:State.OscdimgLabel = "ADK" }
    else {
      $cmd = Get-Command oscdimg.exe -ErrorAction SilentlyContinue
      if (-not $cmd) { Fail "oscdimg.exe not found (install ADK or add to PATH)." }
      $script:State.OscdimgPath = $cmd.Source; $script:State.OscdimgLabel = "PATH"
    }
  }

  V ("Using DISM: {0} ({1})" -f $script:State.DismPath, $script:State.DismLabel)
  V ("Using OSCDIMG: {0} ({1})" -f $script:State.OscdimgPath, $script:State.OscdimgLabel)
}

# ==============================
# Work paths / meta
# ==============================
function Initialize-WorkPaths([string]$FolderPath, [string]$IsoPath, [switch]$UseSystemTemp) {
  $workBase = if ($UseSystemTemp) { [System.IO.Path]::GetTempPath() } else { Join-Path $FolderPath $Config.WorkParentSubfolder }
  New-Item -ItemType Directory -Path $workBase -Force | Out-Null
  $isoBase = [IO.Path]::GetFileNameWithoutExtension($IsoPath)

  $script:State.WorkBase    = $workBase
  $script:State.WorkRoot    = Join-Path $workBase ($Config.WorkPrefix + $isoBase)
  $script:State.IsoRoot     = Join-Path $script:State.WorkRoot $Config.WorkIsoSubdir
  $script:State.InstallRoot = Join-Path $script:State.WorkRoot $Config.WorkInstallSubdir
  $script:State.MountRoot   = Join-Path $script:State.WorkRoot $Config.WorkMountSubdir
  $script:State.LogsRoot    = Join-Path $script:State.WorkRoot $Config.WorkLogsSubdir
  $script:State.ScratchRoot = Join-Path $script:State.WorkRoot $Config.WorkScratchSubdir
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
    "StashedInstallWim=$($script:State.StashedInstallWim)",
    "OutputIsoPath=$($script:State.OutputIsoPath)"
  )
  if (In-AfterPrepDryRun) { Write-Host "[DryRun] Would write meta file." -ForegroundColor Yellow; return }
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
  if ($isos.Count -ne 1) { Fail "Expected exactly 1 input ISO (excluding *$($Config.OutputIsoSuffix)) in '$FolderPath'. Found $($isos.Count)." }
  return $isos[0].FullName
}

function Mount-Iso([string]$IsoPath) {
  if (In-AfterPrepDryRun) { return @{ Drive=$null } }
  $img = Mount-DiskImage -ImagePath $IsoPath -PassThru
  $vol = $img | Get-Volume -ErrorAction SilentlyContinue
  if (-not $vol -or -not $vol.DriveLetter) {
    Dismount-DiskImage -ImagePath $IsoPath -ErrorAction SilentlyContinue | Out-Null
    $img = Mount-DiskImage -ImagePath $IsoPath -PassThru
    $vol = $img | Get-Volume
  }
  if (-not $vol -or -not $vol.DriveLetter) { Fail "Failed to mount ISO or resolve a drive letter." }
  $script:State.IsoWasMounted = $true
  return @{ Drive="$($vol.DriveLetter):" }
}

function Copy-IsoContents([string]$SrcDrive, [string]$DstFolder, [string]$IsoPathForDisplay) {
  if (In-AfterPrepDryRun) { return }

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

  $args = @($srcDisplay, $dstDisplay, "*.*") + $Config.RobocopyArgsBase
  if ($VerbosePreference -ne 'Continue') { $args += $Config.RobocopyArgsQuiet }

  if ($VerbosePreference -ne 'Continue') { robocopy @args *> $logPath } else { robocopy @args }
  $rc = [int]$LASTEXITCODE
  if ($rc -ge 8) { Fail "Robocopy failed with exit code $rc. See log: $logPath" }
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
  if ($LASTEXITCODE -ne 0) { Fail "DISM /Get-WimInfo failed for $InstallPath" }

  $pairs = @()
  $cur = $null
  foreach ($line in $out) {
    if ($line -match '^\s*Index\s*:\s*(\d+)\s*$') { $cur = [int]$Matches[1]; continue }
    if ($cur -ne $null -and $line -match '^\s*Name\s*:\s*(.+)\s*$') {
      $pairs += [pscustomobject]@{ Index=$cur; Name=$Matches[1].Trim() }
      $cur = $null
    }
  }
  if ($pairs.Count -lt 1) { Fail "No indexes found in $InstallPath" }
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
# Indices selection (label/wildcard/regex)
# ==============================
function Normalize-Label([string]$s) {
  if (-not $s) { return "" }
  $t = $s.Trim()
  $t = $t -replace '^(?i)\s*Windows\s+\d+\s+', ''
  $t = $t -replace '\s+', ' '
  return $t.Trim()
}

function Unquote([string]$s) {
  if ($null -eq $s) { return "" }
  $t = $s.Trim()
  if (($t.StartsWith('"') -and $t.EndsWith('"')) -or ($t.StartsWith("'") -and $t.EndsWith("'"))) { return $t.Substring(1, $t.Length-2) }
  return $t
}

function Is-RegexLabelToken([string]$tok) { (Unquote $tok).Trim().StartsWith("re:", [System.StringComparison]::OrdinalIgnoreCase) }
function Is-WildcardLabelToken([string]$tok) {
  $t = Unquote $tok
  if ($t.StartsWith("re:", [System.StringComparison]::OrdinalIgnoreCase)) { return $false }
  return ($t.Contains('*') -or $t.Contains('?'))
}

function Resolve-LabelTokenToIndices {
  param([object[]]$Pairs, [string]$Token)

  $tok = Unquote $Token
  $items = foreach ($p in $Pairs) { [pscustomobject]@{ Index=[int]$p.Index; Norm=(Normalize-Label $p.Name) } }

  if (Is-RegexLabelToken $tok) {
    $pat = $tok.Substring(3).Trim()
    if ($pat -eq "") { return @() }
    try {
      $rx = New-Object System.Text.RegularExpressions.Regex ($pat, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
      return @($items | Where-Object { $rx.IsMatch($_.Norm) } | Select-Object -ExpandProperty Index)
    } catch { return @() }
  }

  if (Is-WildcardLabelToken $tok) {
    return @($items | Where-Object { $_.Norm -like $tok } | Select-Object -ExpandProperty Index)
  }

  $key = (Normalize-Label $tok).ToLowerInvariant()
  return @($items | Where-Object { $_.Norm.ToLowerInvariant() -eq $key } | Select-Object -ExpandProperty Index)
}

function Normalize-IndicesSpec([string]$Spec) {
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

function Parse-IndicesSpec {
  param([string]$Spec, [int]$MaxIndex, [object[]]$Pairs)

  $selected = @()
  $badTokens = @()
  $badIdx = @()
  $badLabels = @()

  $Spec = Normalize-IndicesSpec $Spec
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
# ESD -> WIM
# ==============================
function Ensure-StashedInstallWim {
  param([string]$InstallRoot)

  $wim = Join-Path $InstallRoot 'install.wim'
  $esd = Join-Path $InstallRoot 'install.esd'
  if (Test-Path $wim) { return $wim }
  if (-not (Test-Path $esd)) { return $null }

  Write-Host "Converting INSTALL\install.esd to INSTALL\install.wim..." -ForegroundColor Cyan

  $info = Invoke-DismRead -Args @("/Get-ImageInfo", "/ImageFile:$esd")
  if ($LASTEXITCODE -ne 0) { Fail "DISM failed to read install.esd." }

  $indexes = @()
  foreach ($line in $info) { if ($line -match '^\s*Index\s*:\s*(\d+)\s*$') { $indexes += [int]$Matches[1] } }
  if ($indexes.Count -lt 1) { Fail "Could not parse indexes from install.esd." }

  if (Test-Path $wim) { Remove-Item -Path $wim -Force -ErrorAction SilentlyContinue | Out-Null }

  foreach ($idx in $indexes) {
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
  return $wim
}

# ==============================
# Rebuild install.wim
# ==============================
function Rebuild-InstallWimFromSelection {
  param(
    [string]$StashedWim,
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

  $allSelected = ($SelectedSourceIndices.Count -eq $MaxSourceIndex)
  if ($allSelected) {
    Invoke-Step "Copy full install.wim (all indices) to ISO\sources" { Copy-Item -Path $StashedWim -Destination $dstWim -Force } | Out-Null
    return $dstWim
  }

  Write-Host "Rebuilding ISO\sources\install.wim from selected indices..." -ForegroundColor Cyan
  foreach ($idx in $SelectedSourceIndices) {
    $rc = Invoke-External -FilePath $script:State.DismPath -ArgumentList @(
      "/Export-Image",
      "/SourceImageFile:$StashedWim",
      "/SourceIndex:$idx",
      "/DestinationImageFile:$dstWim",
      "/Compress:max",
      "/CheckIntegrity"
    )
    if ($rc -ne 0) { Fail "DISM Export-Image failed for SourceIndex $idx (exit $rc)." }
  }
  return $dstWim
}

# ==============================
# Post-install script + autounattend
# ==============================
function Write-PostInstallWingetScript([string]$DestPath) {
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
  Invoke-Step "Write PostInstallWinget.ps1" { Set-Content -Path $DestPath -Value $content -Encoding ASCII } | Out-Null
}

function Write-Autounattend {
  param(
    [Parameter(Mandatory=$true)][string]$IsoRoot,
    [Parameter(Mandatory=$true)][bool]$IncludePostInstall
  )

  $path = Join-Path $IsoRoot $Config.AutounattendName
  $driversPath = $Config.DriverFolderUnattendPath

  Invoke-Step "Write autounattend.xml" {
    $ps = "%WINDIR%\Setup\Scripts\$($Config.PostInstallScriptName)"

    $oobeBlock = ""
    if ($IncludePostInstall) {
      $oobeBlock = @"
  <settings pass="oobeSystem">
    <component name="Microsoft-Windows-Shell-Setup"
               processorArchitecture="amd64"
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
"@
    }

    $xml = @"
<?xml version="1.0" encoding="utf-8"?>
<unattend xmlns="urn:schemas-microsoft-com:unattend">
  <settings pass="windowsPE">
    <component name="Microsoft-Windows-PnpCustomizationsWinPE"
               processorArchitecture="amd64"
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
$oobeBlock
</unattend>
"@
    Set-Content -Path $path -Value $xml -Encoding ASCII
  } | Out-Null
}

# ==============================
# Servicing + ISO build
# ==============================
function Add-PackagesFromFolderToMountedImage {
  param([string]$MountDir, [string]$MsuFolder, [string]$ScratchRoot, [string]$LogPath)
  $args = @(
    "/Image:$MountDir",
    "/Add-Package",
    "/PackagePath:$MsuFolder",
    "/ScratchDir:$ScratchRoot",
    "/LogPath:$LogPath"
  )
  $rc = Invoke-External -FilePath $script:State.DismPath -ArgumentList $args
  if ($rc -ne 0) { Fail "DISM Add-Package failed (exit $rc). See log: $LogPath" }
}

function Service-WimIndex {
  param(
    [string]$WimPath, [int]$Index, [string]$IndexName,
    [string]$MsuFolder, [string]$MountRoot,
    [string]$PostInstallScriptSource, [string]$LogsRoot, [string]$ScratchRoot,
    [bool]$DoMsu
  )

  Write-Host ("Processing index {0}: {1}" -f $Index, $IndexName) -ForegroundColor Cyan

  $mountDir = Join-Path $MountRoot ("idx_{0}" -f $Index)
  $mountLog = Join-Path $LogsRoot ("dism_mount_idx{0}.log" -f $Index)

  Invoke-Step ("Prepare mount dir for index $Index") {
    if (Test-Path $mountDir) { Remove-Item $mountDir -Recurse -Force -ErrorAction SilentlyContinue | Out-Null }
    New-Item -ItemType Directory -Path $mountDir -Force | Out-Null
  } | Out-Null

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
    $setupScripts = Join-Path $mountDir 'Windows\Setup\Scripts'
    Invoke-Step ("Create setup scripts dir (idx $Index)") { New-Item -ItemType Directory -Path $setupScripts -Force | Out-Null } | Out-Null
    Invoke-Step ("Copy post-install script (idx $Index)") {
      Copy-Item -Path $PostInstallScriptSource -Destination (Join-Path $setupScripts $Config.PostInstallScriptName) -Force
    } | Out-Null

    if ($DoMsu) {
      $pkgLog = Join-Path $LogsRoot ("dism_addpackage_idx{0}.log" -f $Index)
      Add-PackagesFromFolderToMountedImage -MountDir $mountDir -MsuFolder $MsuFolder -ScratchRoot $ScratchRoot -LogPath $pkgLog
    }
  }
  finally {
    $rc2 = Invoke-External -FilePath $script:State.DismPath -ArgumentList @(
      "/Unmount-Image",
      "/MountDir:$mountDir",
      "/Commit"
    )
    if ($rc2 -ne 0) { Fail "Failed to unmount/commit index $Index (exit $rc2)." }
  }
}

function Build-Iso([string]$IsoRoot, [string]$OutputIso) {
  $etfs = Join-Path $IsoRoot $Config.BootFileBIOS
  $efis = Join-Path $IsoRoot $Config.BootFileUEFI
  if (-not (Test-Path $etfs)) { Fail "Missing BIOS boot file: $etfs" }
  if (-not (Test-Path $efis)) { Fail "Missing UEFI boot file: $efis" }

  $bootdata = "2#p0,e,b$etfs#pEF,e,b$efis"
  $args = @() + $Config.OscdimgFsArgs + @("-l$($Config.IsoVolumeLabel)", "-bootdata:$bootdata", $IsoRoot, $OutputIso)

  Write-Host "Building ISO: $OutputIso" -ForegroundColor Green
  $rc = Invoke-External -FilePath $script:State.OscdimgPath -ArgumentList $args
  if ($rc -ne 0) { Fail "oscdimg failed with exit code $rc" }
}

function Cleanup-Hardened {
  Stop-TrackedChildren
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
$UpdateISO = $false
$ShowIndices = $false
$SelectHome = $false
$SelectPro = $false
$IndicesSpec = $null

for ($i = 0; $i -lt $args.Count; $i++) {
  $a = $args[$i]
  switch -Regex ($a) {
    '^(?:-h|-help|-\\?|/\\?)$' { $ShowHelp = $true; continue }
    '^(?:-UseSystemTemp)$'     { $UseSystemTemp = $true; continue }
    '^(?:-Verbose|-v)$'        { $VerboseSwitch = $true; continue }
    '^(?:-DryRun)$'            { $script:DryRun = $true; continue }
    '^(?:-UseADK)$'            { $UseADK = $true; continue }
    '^(?:-UseSystem)$'         { $UseSystem = $true; continue }
    '^(?:-CleanWork)$'         { $CleanWork = $true; continue }
    '^(?:-UpdateISO)$'         { $UpdateISO = $true; continue }
    '^(?:-ShowIndices)$'       { $ShowIndices = $true; continue }
    '^(?:-Show)$'              { $ShowIndices = $true; continue }
    '^(?:-Home)$'              { $SelectHome = $true; continue }
    '^(?:-Pro)$'               { $SelectPro = $true; continue }
    '^(?:-Indices)$' {
      if ($i + 1 -ge $args.Count) { Fail "-Indices requires one or more selector tokens" }
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
    '^-{1,2}.*' { Fail "Unknown switch: $a" }
    default {
      if ($FolderArg) { Fail "Only one folder argument allowed. Extra value: $a" }
      $FolderArg = $a
    }
  }
}

if ($VerboseSwitch) { $VerbosePreference = 'Continue' }
if ($ShowHelp) { Show-HelpAndExit }
if (-not $FolderArg) { $FolderArg = (Get-Location).Path }

# ==============================
# MAIN
# ==============================
Assert-Admin
Resolve-Tools -ExplicitDism $ExplicitDism -ExplicitOscdimg $ExplicitOscdimg -UseADK:($UseADK) -UseSystem:($UseSystem)

try {
  $folderPath = (Resolve-Path -LiteralPath $FolderArg -ErrorAction SilentlyContinue).Path
  if (-not $folderPath) { Fail "Folder not found: $FolderArg" }

  # Resolve ISO path (new run) or meta (update run)
  if (-not $UpdateISO) {
    $script:State.IsoPath = Get-InputIso -FolderPath $folderPath
    Initialize-WorkPaths -FolderPath $folderPath -IsoPath $script:State.IsoPath -UseSystemTemp:($UseSystemTemp)
  } else {
    $workBase = if ($UseSystemTemp) { [System.IO.Path]::GetTempPath() } else { Join-Path $folderPath $Config.WorkParentSubfolder }
    $metaFiles = Get-ChildItem -LiteralPath $workBase -Filter $Config.MetaFileName -Recurse -File -ErrorAction SilentlyContinue
    if ($metaFiles.Count -ne 1) { Fail "UpdateISO requires exactly one meta file under $workBase. Found $($metaFiles.Count)." }
    $meta = Read-Meta -MetaPath $metaFiles[0].FullName
    if (-not $meta) { Fail "Failed to read meta file: $($metaFiles[0].FullName)" }

    $script:State.IsoPath = $meta["ISOPath"]
    $script:State.IsoBaseName = $meta["IsoBaseName"]
    $script:State.WorkRoot = $meta["WorkRoot"]
    $script:State.IsoRoot = $meta["IsoRoot"]
    $script:State.InstallRoot = $meta["InstallRoot"]
    $script:State.MountRoot = $meta["MountRoot"]
    $script:State.LogsRoot = $meta["LogsRoot"]
    $script:State.ScratchRoot = $meta["ScratchRoot"]
    $script:State.StashedInstallWim = $meta["StashedInstallWim"]
    $script:State.OutputIsoPath = $meta["OutputIsoPath"]
    $script:State.MetaPath = $metaFiles[0].FullName
  }

  if (-not $script:State.OutputIsoPath) {
    $baseName = [IO.Path]::GetFileNameWithoutExtension($script:State.IsoPath)
    $script:State.OutputIsoPath = Join-Path $folderPath ($baseName + $Config.OutputIsoSuffix)
  }

  if ($CleanWork -and (Test-Path $script:State.WorkRoot)) {
    Remove-Item -Path $script:State.WorkRoot -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
  }

  # PREP always runs
  $script:DryRunPhase = 'prep'
  New-Item -ItemType Directory -Path $script:State.WorkRoot -Force | Out-Null
  New-Item -ItemType Directory -Path $script:State.IsoRoot, $script:State.InstallRoot, $script:State.MountRoot, $script:State.LogsRoot, $script:State.ScratchRoot -Force | Out-Null

  # ISO extract (only if not UpdateISO)
  if (-not $UpdateISO) {
    $needsCopy = -not (Get-ChildItem -LiteralPath $script:State.IsoRoot -Force -ErrorAction SilentlyContinue | Select-Object -First 1)
    if ($needsCopy) {
      $isoMount = Mount-Iso $script:State.IsoPath
      Copy-IsoContents -SrcDrive $isoMount.Drive -DstFolder $script:State.IsoRoot -IsoPathForDisplay $script:State.IsoPath
    }
  }

  # Ensure Drivers folder exists in ISO root (so it exists in the final ISO)
  $driversFolderOnIso = Join-Path $script:State.IsoRoot $Config.DriverFolderName
  New-Item -ItemType Directory -Path $driversFolderOnIso -Force | Out-Null
  $driverReadme = Join-Path $driversFolderOnIso "README.txt"
  if (-not (Test-Path $driverReadme)) {
    Set-Content -Path $driverReadme -Encoding ASCII -Value @(
      "Place INF-based drivers under this folder (subfolders allowed).",
      "Unattend path: $($Config.DriverFolderUnattendPath)"
    )
  }

  $isoSources = Join-Path $script:State.IsoRoot 'sources'
  if (-not (Test-Path $isoSources)) { Fail "ISO sources directory not found: $isoSources" }

  # Move install.* to INSTALL if needed
  $stashWim = Join-Path $script:State.InstallRoot 'install.wim'
  $stashEsd = Join-Path $script:State.InstallRoot 'install.esd'
  if (-not (Test-Path $stashWim) -and -not (Test-Path $stashEsd)) {
    $srcInstall = Get-InstallImagePathFromRoot -Root $script:State.IsoRoot
    if (-not $srcInstall) { Fail "Could not find sources\install.wim or sources\install.esd in ISO tree." }
    Move-Item -Path $srcInstall -Destination (Join-Path $script:State.InstallRoot (Split-Path -Leaf $srcInstall)) -Force
  }

  if (-not $script:State.StashedInstallWim) {
    $script:State.StashedInstallWim = Ensure-StashedInstallWim -InstallRoot $script:State.InstallRoot
  }
  if (-not $script:State.StashedInstallWim) { Fail "INSTALL stash does not contain install.wim or install.esd." }

  # AFTER PREP
  $script:DryRunPhase = 'afterprep'

  if ($ShowIndices) { Show-Indices -InstallPath $script:State.StashedInstallWim; return }

  $pairs = Get-WimPairs -InstallPath $script:State.StashedInstallWim
  $maxIndex = ($pairs | Measure-Object -Property Index -Maximum).Maximum
  $nameMap = Get-IndexNameMap -Pairs $pairs

  $explicitSelectionUsed = ($SelectHome -or $SelectPro -or ($IndicesSpec -and ([string]$IndicesSpec).Trim() -ne ""))

  # UpdateISO semantics:
  # - If UpdateISO and no explicit indices: selected empty, and do NOT rebuild, do NOT service, do NOT MSU
  $selected = @()
  $doRebuild = $true
  $doService = $true
  $doMsu = $true

  if ($UpdateISO -and -not $explicitSelectionUsed) {
    $selected = @()
    $doRebuild = $false
    $doService = $false
    $doMsu = $false
    Write-Host "[UpdateISO] No indices specified; skipping rebuild, servicing, and MSU handling." -ForegroundColor Yellow
  } else {
    if (-not $explicitSelectionUsed -and -not $UpdateISO) {
      $selected = 1..$maxIndex
    } else {
      if ($SelectHome) { $selected += (Resolve-LabelTokenToIndices -Pairs $pairs -Token "Home") }
      if ($SelectPro)  { $selected += (Resolve-LabelTokenToIndices -Pairs $pairs -Token "Pro") }

      if ($IndicesSpec -and ([string]$IndicesSpec).Trim() -ne "") {
        $parsed = Parse-IndicesSpec -Spec ([string]$IndicesSpec) -MaxIndex $maxIndex -Pairs $pairs
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
      if ($selected.Count -lt 1) { Fail "Selection resulted in an empty set." }
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
    if ($doRebuild) { Write-Host "[DryRun] Would rebuild ISO\sources\install.wim." -ForegroundColor Yellow }
    else { Write-Host "[DryRun] Would not rebuild ISO\sources\install.wim." -ForegroundColor Yellow }
    if ($doService) { Write-Host "[DryRun] Would service image(s)." -ForegroundColor Yellow }
    else { Write-Host "[DryRun] Would not service image(s)." -ForegroundColor Yellow }
    if ($doMsu) { Write-Host "[DryRun] Would handle/apply MSUs." -ForegroundColor Yellow }
    else { Write-Host "[DryRun] Would not handle/apply MSUs." -ForegroundColor Yellow }
    Write-Host ("[DryRun] Would write autounattend.xml with driver path: {0}" -f $Config.DriverFolderUnattendPath) -ForegroundColor Yellow
    Write-Host ("[DryRun] Would build ISO: {0}" -f $script:State.OutputIsoPath) -ForegroundColor Yellow
    return
  }

  # Ensure autounattend.xml exists (always includes Drivers path; only includes postinstall if servicing)
  Write-Autounattend -IsoRoot $script:State.IsoRoot -IncludePostInstall:$doService

  # If UpdateISO + no indices, do not service, do not MSU, do not rebuild; just build ISO from existing tree
  if (-not $doService) {
    Build-Iso -IsoRoot $script:State.IsoRoot -OutputIso $script:State.OutputIsoPath
    Write-Meta
    Write-Host "SUCCESS: Created bundled ISO:" -ForegroundColor Green
    Write-Host ("  " + $script:State.OutputIsoPath) -ForegroundColor Green
    Write-Host "Work folder preserved at:" -ForegroundColor Yellow
    Write-Host ("  " + $script:State.WorkRoot) -ForegroundColor Yellow
    return
  }

  # Determine WIM to service
  $wimToService = $null
  if ($doRebuild) {
    $wimToService = Rebuild-InstallWimFromSelection -StashedWim $script:State.StashedInstallWim -IsoSourcesDir $isoSources -SelectedSourceIndices $selected -MaxSourceIndex $maxIndex
  } else {
    $wimToService = Join-Path $isoSources 'install.wim'
    if (-not (Test-Path $wimToService)) {
      Fail "Skipped rebuild, but ISO\sources\install.wim is missing. Specify indices to force rebuild."
    }
  }

  # Determine MSU folder and whether to apply
  $msuFolder = $folderPath
  $msuFiles = @()
  if ($doMsu) {
    $msuFiles = @(Get-ChildItem -LiteralPath $folderPath -Filter "*.msu" -File -ErrorAction SilentlyContinue)
    V ("MSU files found (top-level only): {0}" -f $msuFiles.Count)
    if ($msuFiles.Count -lt 1) {
      # No download here by design (you asked not to merge/apply MSUs when UpdateISO has no indices).
      # If you want catalog-download back for non-UpdateISO runs, re-add your MSCatalogLTS logic here.
      $doMsu = $false
      V "No MSUs present; skipping MSU application."
    }
  }

  # Prepare postinstall script (only relevant if servicing)
  $postInstallTemp = Join-Path $script:State.MountRoot $Config.PostInstallScriptName
  Write-PostInstallWingetScript -DestPath $postInstallTemp

  # Service all indices in chosen install.wim
  $rebuiltPairs = Get-WimPairs -InstallPath $wimToService
  $rebuiltNameMap = Get-IndexNameMap -Pairs $rebuiltPairs
  $serviceIndexes = @($rebuiltPairs | Select-Object -ExpandProperty Index)

  Write-Host "Servicing install.wim (index : name):" -ForegroundColor Cyan
  foreach ($idx in $serviceIndexes) {
    $nm = $rebuiltNameMap[[int]$idx]; if (-not $nm) { $nm = "<unknown>" }
    Write-Host ("  {0}: {1}" -f $idx, $nm) -ForegroundColor Cyan
  }

  foreach ($idx in $serviceIndexes) {
    $nm = $rebuiltNameMap[[int]$idx]; if (-not $nm) { $nm = "<unknown>" }
    Service-WimIndex -WimPath $wimToService -Index $idx -IndexName $nm -MsuFolder $msuFolder -MountRoot $script:State.MountRoot -PostInstallScriptSource $postInstallTemp -LogsRoot $script:State.LogsRoot -ScratchRoot $script:State.ScratchRoot -DoMsu:$doMsu
  }

  Build-Iso -IsoRoot $script:State.IsoRoot -OutputIso $script:State.OutputIsoPath
  Write-Meta

  Write-Host "SUCCESS: Created bundled ISO:" -ForegroundColor Green
  Write-Host ("  " + $script:State.OutputIsoPath) -ForegroundColor Green
  Write-Host "Work folder preserved at:" -ForegroundColor Yellow
  Write-Host ("  " + $script:State.WorkRoot) -ForegroundColor Yellow
}
finally {
  Cleanup-Hardened
}
