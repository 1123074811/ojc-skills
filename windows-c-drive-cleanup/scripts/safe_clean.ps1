<#
.SYNOPSIS
  Safe cleanup of temp files, updater leftovers, regenerable caches.
.EXAMPLE
  pwsh -NoProfile -File .\safe_clean.ps1 -WhatIf
  pwsh -NoProfile -File .\safe_clean.ps1 -UserName MECHREVO -CleanDownloadsInstallers
#>
[CmdletBinding()]
param(
  [string]$UserName = $env:USERNAME,
  [switch]$WhatIf,
  [switch]$Execute,
  [string[]]$ApprovedPath = @(),
  [switch]$CleanDownloadsInstallers,
  [switch]$EmptyRecycleBin,
  [switch]$CleanClaudeVmBundles
)

$ErrorActionPreference = 'SilentlyContinue'
$userProfile = Join-Path 'C:\Users' $UserName
if (-not (Test-Path $userProfile)) { $userProfile = $env:USERPROFILE }
$local = Join-Path $userProfile 'AppData\Local'
$roaming = Join-Path $userProfile 'AppData\Roaming'
$PreviewOnly = $WhatIf -or -not $Execute
$approved = @($ApprovedPath | ForEach-Object { [IO.Path]::GetFullPath($_).TrimEnd('\') })

function Get-SizeBytes([string]$Path) {
  if (-not (Test-Path -LiteralPath $Path)) { return [int64]0 }
  $item = Get-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
  if (-not $item) { return [int64]0 }
  if (-not $item.PSIsContainer) { return [int64]$item.Length }
  $sum = (Get-ChildItem -LiteralPath $Path -Force -Recurse -File -ErrorAction SilentlyContinue |
    Measure-Object Length -Sum).Sum
  if ($null -eq $sum) { return [int64]0 }
  return [int64]$sum
}

function Clear-DirContents([string]$Path) {
  if (-not (Test-Path -LiteralPath $Path)) { return 'skip-missing' }
  if ($PreviewOnly) { return 'preview' }
  Get-ChildItem -LiteralPath $Path -Force -ErrorAction SilentlyContinue | ForEach-Object {
    Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
  }
  return 'cleaned'
}

$before = (Get-PSDrive C).Free
Write-Host ("C free before: {0:N2} GB" -f ($before / 1GB))
if ($PreviewOnly) { Write-Host 'PREVIEW ONLY — user approval and -Execute are required for deletions' -ForegroundColor Yellow }

$targets = @(
  (Join-Path $local 'Temp'),
  'C:\Windows\Temp',
  'C:\Temp',
  'C:\tmp',
  (Join-Path $local 'npm-cache'),
  (Join-Path $local 'pip'),
  (Join-Path $userProfile '.cache'),
  (Join-Path $userProfile '.npm'),
  (Join-Path $roaming 'Code\CachedExtensionVSIXs'),
  (Join-Path $roaming 'Code\Cache'),
  (Join-Path $roaming 'Code\CachedData'),
  (Join-Path $roaming 'Code\logs'),
  (Join-Path $local 'Microsoft\Windows\INetCache'),
  (Join-Path $local 'Microsoft\PowerToys\Updates'),
  (Join-Path $local '@genieworkbuddy-desktop-updater'),
  (Join-Path $local 'antigravity-updater'),
  (Join-Path $local 'cherrystudio-updater'),
  'C:\Windows\SoftwareDistribution\Download'
)

# Dynamic updater dirs
Get-ChildItem -LiteralPath $local -Force -Directory -ErrorAction SilentlyContinue |
  Where-Object { $_.Name -match 'updater|Update' } |
  ForEach-Object { $targets += $_.FullName }

$targets = $targets | Select-Object -Unique
$report = @()

foreach ($p in $targets) {
  $beforeSize = Get-SizeBytes $p
  $normalized = [IO.Path]::GetFullPath($p).TrimEnd('\')
  $status = if ($Execute -and $normalized -notin $approved) { 'not-approved' } else { Clear-DirContents $p }
  $afterSize = Get-SizeBytes $p
  $freed = [math]::Max(0, $beforeSize - $afterSize)
  $report += [PSCustomObject]@{
    Path     = $p
    BeforeMB = [math]::Round($beforeSize / 1MB, 1)
    FreedMB  = [math]::Round($freed / 1MB, 1)
    Status   = $status
  }
}

if ($Execute -and $approved.Count -eq 0 -and -not ($CleanDownloadsInstallers -or $EmptyRecycleBin -or $CleanClaudeVmBundles)) {
  throw 'No approved action. Pass exact user-approved paths with -ApprovedPath, or use an explicitly approved optional action switch.'
}

if ($CleanDownloadsInstallers) {
  $dl = Join-Path $userProfile 'Downloads'
  Get-ChildItem -LiteralPath $dl -Force -File -ErrorAction SilentlyContinue |
    Where-Object {
      $_.Extension -match '\.(exe|msi|msix)$' -and (
        $_.Name -match 'Setup|Installer|UserSetup|Netdisk|Qoder|Baidu|x64|x86'
      )
    } | ForEach-Object {
      $sz = $_.Length
      if (-not $PreviewOnly) {
        Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
      }
      $ok = $PreviewOnly -or (-not (Test-Path -LiteralPath $_.FullName))
      $report += [PSCustomObject]@{
        Path     = $_.FullName
        BeforeMB = [math]::Round($sz / 1MB, 1)
        FreedMB  = if ($ok -and -not $PreviewOnly) { [math]::Round($sz / 1MB, 1) } else { 0 }
        Status   = if ($PreviewOnly) { 'preview' } elseif ($ok) { 'deleted' } else { 'locked' }
      }
    }
}

if ($CleanClaudeVmBundles) {
  $pkgRoot = Join-Path $local 'Packages'
  Get-ChildItem -LiteralPath $pkgRoot -Force -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -like 'Claude_*' } |
    ForEach-Object {
      $vm = Join-Path $_.FullName 'LocalCache\Roaming\Claude\vm_bundles'
      if (Test-Path -LiteralPath $vm) {
        $beforeSize = Get-SizeBytes $vm
        $status = Clear-DirContents $vm
        $afterSize = Get-SizeBytes $vm
        $report += [PSCustomObject]@{
          Path     = $vm
          BeforeMB = [math]::Round($beforeSize / 1MB, 1)
          FreedMB  = [math]::Round([math]::Max(0, $beforeSize - $afterSize) / 1MB, 1)
          Status   = $status
        }
      }
    }
}

if ($EmptyRecycleBin -and -not $PreviewOnly) {
  try {
    Clear-RecycleBin -Force -ErrorAction SilentlyContinue
    $report += [PSCustomObject]@{ Path = 'RecycleBin'; BeforeMB = $null; FreedMB = $null; Status = 'emptied' }
  } catch {
    $report += [PSCustomObject]@{ Path = 'RecycleBin'; BeforeMB = $null; FreedMB = $null; Status = 'skip' }
  }
}
elseif ($EmptyRecycleBin) {
  $report += [PSCustomObject]@{ Path = 'RecycleBin'; BeforeMB = $null; FreedMB = 0; Status = 'preview' }
}

$after = (Get-PSDrive C).Free
Write-Host "`n=== Cleanup report ===" -ForegroundColor Cyan
$report | Sort-Object FreedMB -Descending | Format-Table -AutoSize
Write-Host ("C free after: {0:N2} GB" -f ($after / 1GB))
Write-Host ("Drive free delta: {0:N2} GB" -f (($after - $before) / 1GB))
