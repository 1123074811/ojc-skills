<#
.SYNOPSIS
  Audit C: drive large folders and files for cleanup/migration planning.
.EXAMPLE
  pwsh -NoProfile -File .\scan_c_drive.ps1 -UserName MECHREVO
#>
[CmdletBinding()]
param(
  [string]$UserName = $env:USERNAME,
  [string]$Drive = 'C',
  [int]$MinFileMB = 100,
  [int]$TopFiles = 60
)

$ErrorActionPreference = 'SilentlyContinue'
  $userProfile = Join-Path (Join-Path ($Drive + ':') 'Users') $UserName
if (-not (Test-Path -LiteralPath $userProfile)) {
  $userProfile = $env:USERPROFILE
}

function Get-DirSizeBytes([string]$Path) {
  if (-not (Test-Path -LiteralPath $Path)) { return $null }
  $sum = (Get-ChildItem -LiteralPath $Path -Force -Recurse -File -ErrorAction SilentlyContinue |
    Measure-Object Length -Sum).Sum
  if ($null -eq $sum) { return [int64]0 }
  return [int64]$sum
}

function Format-GB([int64]$Bytes) {
  return [math]::Round($Bytes / 1GB, 2)
}

Write-Host "=== Drive free space ===" -ForegroundColor Cyan
Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Used -or $_.Free } | ForEach-Object {
  [PSCustomObject]@{
    Name    = $_.Name
    UsedGB  = [math]::Round($_.Used / 1GB, 2)
    FreeGB  = [math]::Round($_.Free / 1GB, 2)
    TotalGB = [math]::Round(($_.Used + $_.Free) / 1GB, 2)
  }
} | Format-Table -AutoSize

Write-Host "=== $Drive`:\ top-level ===" -ForegroundColor Cyan
$top = Get-ChildItem -LiteralPath "$Drive`:" -Force -ErrorAction SilentlyContinue | ForEach-Object {
  if ($_.PSIsContainer) {
    $b = Get-DirSizeBytes $_.FullName
    [PSCustomObject]@{ Name = $_.Name; Type = 'Dir'; GB = (Format-GB $b); Path = $_.FullName }
  } else {
    [PSCustomObject]@{ Name = $_.Name; Type = 'File'; GB = (Format-GB ([int64]$_.Length)); Path = $_.FullName }
  }
}
$top | Sort-Object GB -Descending | Select-Object -First 25 | Format-Table -AutoSize

Write-Host "=== User profile top ===" -ForegroundColor Cyan
if (Test-Path -LiteralPath $userProfile) {
  $uprofile = Get-ChildItem -LiteralPath $userProfile -Force -ErrorAction SilentlyContinue | ForEach-Object {
    if ($_.PSIsContainer) {
      $b = Get-DirSizeBytes $_.FullName
      [PSCustomObject]@{ Name = $_.Name; GB = (Format-GB $b); Path = $_.FullName }
    } else {
      [PSCustomObject]@{ Name = $_.Name; GB = (Format-GB ([int64]$_.Length)); Path = $_.FullName }
    }
  }
  $uprofile | Sort-Object GB -Descending | Select-Object -First 30 | Format-Table -AutoSize
}

Write-Host "=== AppData Local top ===" -ForegroundColor Cyan
$local = Join-Path $userProfile 'AppData\Local'
if (Test-Path $local) {
  $rows = Get-ChildItem -LiteralPath $local -Force -Directory -ErrorAction SilentlyContinue | ForEach-Object {
    $b = Get-DirSizeBytes $_.FullName
    [PSCustomObject]@{ Name = $_.Name; GB = (Format-GB $b); Path = $_.FullName }
  }
  $rows | Sort-Object GB -Descending | Select-Object -First 25 | Format-Table -AutoSize
}

Write-Host "=== AppData Roaming top ===" -ForegroundColor Cyan
$roaming = Join-Path $userProfile 'AppData\Roaming'
if (Test-Path $roaming) {
  $rows = Get-ChildItem -LiteralPath $roaming -Force -Directory -ErrorAction SilentlyContinue | ForEach-Object {
    $b = Get-DirSizeBytes $_.FullName
    [PSCustomObject]@{ Name = $_.Name; GB = (Format-GB $b); Path = $_.FullName }
  }
  $rows | Sort-Object GB -Descending | Select-Object -First 20 | Format-Table -AutoSize
}

Write-Host "=== Known candidate paths ===" -ForegroundColor Cyan
$candidates = @(
  (Join-Path $userProfile '.gradle'),
  (Join-Path $userProfile '.m2'),
  (Join-Path $userProfile '.jdks'),
  (Join-Path $userProfile '.cache'),
  (Join-Path $userProfile 'miniconda3'),
  (Join-Path $userProfile 'develop'),
  (Join-Path $userProfile 'Downloads'),
  (Join-Path $local 'Docker'),
  (Join-Path $local 'Android'),
  (Join-Path $local 'Temp'),
  (Join-Path $local 'npm-cache'),
  (Join-Path $local 'pip'),
  (Join-Path $local 'Pub'),
  (Join-Path $local 'Packages'),
  (Join-Path $roaming 'npm'),
  (Join-Path $roaming 'Code\CachedExtensionVSIXs'),
  "$Drive`:\pagefile.sys",
  "$Drive`:\hiberfil.sys",
  "$Drive`:\mongodb",
  "$Drive`:\Windows\SoftwareDistribution\Download"
)

$candRows = foreach ($p in $candidates) {
  if (Test-Path -LiteralPath $p) {
    $item = Get-Item -LiteralPath $p -Force
    $bytes = if ($item.PSIsContainer) { Get-DirSizeBytes $p } else { [int64]$item.Length }
    [PSCustomObject]@{
      GB       = Format-GB $bytes
      LinkType = $item.LinkType
      Target   = ($item.Target -join ';')
      Path     = $p
    }
  }
}
$candRows | Sort-Object GB -Descending | Format-Table -AutoSize

Write-Host "=== Large files (>= ${MinFileMB}MB) ===" -ForegroundColor Cyan
$roots = @($userProfile, "$Drive`:\ProgramData", "$Drive`:\mongodb", "$Drive`:\tmp", "$Drive`:\Temp") |
  Where-Object { $_ -and (Test-Path $_) }
$minBytes = [int64]$MinFileMB * 1MB
Get-ChildItem -LiteralPath $roots -Force -Recurse -File -ErrorAction SilentlyContinue |
  Where-Object { $_.Length -ge $minBytes } |
  Sort-Object Length -Descending |
  Select-Object -First $TopFiles @{N='GB';E={[math]::Round($_.Length/1GB,2)}}, @{N='MB';E={[math]::Round($_.Length/1MB,0)}}, FullName |
  Format-Table -AutoSize

Write-Host "=== VHDX files under user Local ===" -ForegroundColor Cyan
Get-ChildItem -LiteralPath $local -Recurse -Include *.vhdx,*.vhd -File -ErrorAction SilentlyContinue |
  Select-Object @{N='GB';E={[math]::Round($_.Length/1GB,2)}}, FullName |
  Sort-Object GB -Descending |
  Select-Object -First 20 |
  Format-Table -AutoSize

Write-Host "=== TRUE on-C sizes (robocopy /XJ, junction-excluded) ===" -ForegroundColor Cyan
function Get-TrueSizeGB([string]$src) {
  if (-not (Test-Path -LiteralPath $src)) { return 0 }
  $out = (& robocopy $src (Join-Path $env:TEMP '_robodummy') /L /E /XJ /NFL /NDL /NJH /BYTES /R:0 /W:0) | Out-String
  $m = [regex]::Match($out, 'Bytes\s*:\s*(\d+)')
  if ($m.Success) { return [math]::Round([double]$m.Groups[1].Value / 1GB, 2) }
  return 0
}
$trueTargets = @($userProfile, (Join-Path $userProfile 'AppData'), $local, $roaming)
$trueRows = foreach ($t in $trueTargets) {
  [PSCustomObject]@{ TrueGB = (Get-TrueSizeGB $t); Path = $t }
}
$trueRows | Format-Table -AutoSize

Write-Host "=== Other profiles under C:\Users ===" -ForegroundColor Cyan
$otherProfiles = Get-ChildItem -LiteralPath 'C:\Users' -Force -Directory -ErrorAction SilentlyContinue |
  Where-Object { $_.Name -ne $UserName -and $_.Name -notin @('Public', 'Default', 'Default User', 'All Users') }
$otherRows = foreach ($op in $otherProfiles) {
  [PSCustomObject]@{ TrueGB = (Get-TrueSizeGB $op.FullName); Path = $op.FullName }
}
$otherRows | Sort-Object TrueGB -Descending | Format-Table -AutoSize

Write-Host "Scan complete." -ForegroundColor Green
