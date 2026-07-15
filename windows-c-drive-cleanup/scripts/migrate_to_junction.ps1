<#
.SYNOPSIS
  Migrate a folder to another drive via robocopy + directory junction, or batch default DevCache set.
.EXAMPLE
  pwsh -NoProfile -File .\migrate_to_junction.ps1 -Source C:\Users\me\.gradle -Dest D:\DevCache\gradle -Label gradle
  pwsh -NoProfile -File .\migrate_to_junction.ps1 -UseDefaultDevCache -TargetDrive D -UserName MECHREVO
#>
[CmdletBinding()]
param(
  [string]$Source,
  [string]$Dest,
  [string]$Label = 'item',
  [switch]$UseDefaultDevCache,
  [string]$TargetDrive = 'D',
  [string]$UserName = $env:USERNAME,
  [switch]$ConfigureEnv,
  [switch]$Execute
)

$ErrorActionPreference = 'Continue'

function Get-DirSizeGB([string]$Path) {
  if (-not (Test-Path -LiteralPath $Path)) { return 0 }
  $sum = (Get-ChildItem -LiteralPath $Path -Force -Recurse -File -ErrorAction SilentlyContinue |
    Measure-Object Length -Sum).Sum
  if ($null -eq $sum) { return 0 }
  return [math]::Round($sum / 1GB, 2)
}

function Migrate-ToJunction {
  param(
    [string]$Source,
    [string]$Dest,
    [string]$Label
  )

  Write-Host ""
  Write-Host "==== Migrating $Label ====" -ForegroundColor Cyan
  Write-Host "Source: $Source"
  Write-Host "Dest  : $Dest"

  if (-not (Test-Path -LiteralPath $Source)) {
    Write-Host 'SKIP: source missing'
    return [PSCustomObject]@{ Label = $Label; Status = 'skip-missing'; GB = 0; Source = $Source; Dest = $Dest }
  }

  $srcItem = Get-Item -LiteralPath $Source -Force
  if ($srcItem.LinkType -eq 'Junction' -or $srcItem.LinkType -eq 'SymbolicLink') {
    Write-Host ("SKIP: already link -> {0}" -f ($srcItem.Target -join ';'))
    return [PSCustomObject]@{
      Label  = $Label
      Status = 'already-linked'
      GB     = (Get-DirSizeGB $Source)
      Source = $Source
      Dest   = ($srcItem.Target -join ';')
    }
  }

  $size = Get-DirSizeGB $Source
  Write-Host ("Size: {0} GB" -f $size)

  $destRoot = Split-Path -Qualifier $Dest
  $destDriveLetter = $destRoot.TrimEnd(':')
  $free = (Get-PSDrive $destDriveLetter -ErrorAction SilentlyContinue).Free
  if ($null -ne $free -and $free -lt ($size * 1GB * 1.05)) {
    Write-Host 'FAIL: not enough free space on target drive'
    return [PSCustomObject]@{ Label = $Label; Status = 'no-space'; GB = $size; Source = $Source; Dest = $Dest }
  }

  if (-not $Execute) {
    Write-Host 'PLAN ONLY: user approval and -Execute are required for migration' -ForegroundColor Yellow
    return [PSCustomObject]@{ Label = $Label; Status = 'awaiting-approval'; GB = $size; Source = $Source; Dest = $Dest }
  }

  $destParent = Split-Path -Parent $Dest
  if (-not (Test-Path -LiteralPath $destParent)) {
    New-Item -ItemType Directory -Force -Path $destParent | Out-Null
  }
  if (-not (Test-Path -LiteralPath $Dest)) {
    New-Item -ItemType Directory -Force -Path $Dest | Out-Null
  }

  & robocopy $Source $Dest /E /COPY:DAT /R:2 /W:1 /XJ /MT:8 /NFL /NDL /NP /NJH /NJS | Out-Null
  $rc = $LASTEXITCODE
  Write-Host ("robocopy exit: {0} (0-7 ok)" -f $rc)
  $destAfter = Get-DirSizeGB $Dest
  Write-Host ("Dest size after copy: {0} GB" -f $destAfter)

  if ($rc -ge 8) {
    Write-Host 'FAIL: robocopy reported an error; source remains unchanged'
    return [PSCustomObject]@{ Label = $Label; Status = 'robocopy-fail'; GB = $size; Source = $Source; Dest = $Dest }
  }

  $backup = $Source + '.pre-migrate-bak'
  if (Test-Path -LiteralPath $backup) {
    Write-Host 'STOP: previous backup exists; inspect it before retrying'
    return [PSCustomObject]@{ Label = $Label; Status = 'backup-exists'; GB = $size; Source = $Source; Dest = $Dest }
  }

  $renamed = $false
  try {
    Rename-Item -LiteralPath $Source -NewName (Split-Path -Leaf $backup) -ErrorAction Stop
    $renamed = $true
  } catch {
    Write-Host ("Rename failed: {0}" -f $_.Exception.Message)
    cmd /c "move /Y `"$Source`" `"$backup`"" | Out-Host
    if ((Test-Path -LiteralPath $backup) -and -not (Test-Path -LiteralPath $Source)) { $renamed = $true }
  }

  if (-not $renamed) {
    Write-Host 'STOP: source is locked; source and copied destination are both preserved'
    return [PSCustomObject]@{ Label = $Label; Status = 'rename-locked'; GB = $size; Source = $Source; Dest = $Dest }
  }

  cmd /c mklink /J "$Source" "$Dest" | Out-Host
  if (-not (Test-Path -LiteralPath $Source)) {
    Rename-Item -LiteralPath $backup -NewName (Split-Path -Leaf $Source) -ErrorAction SilentlyContinue
    return [PSCustomObject]@{ Label = $Label; Status = 'junction-fail'; GB = $size; Source = $Source; Dest = $Dest }
  }

  $link = Get-Item -LiteralPath $Source -Force
  if ($link.LinkType -ne 'Junction') {
    Write-Host 'FAIL: source path exists but is not a junction; backup is preserved'
    return [PSCustomObject]@{ Label = $Label; Status = 'junction-invalid'; GB = $size; Source = $Source; Dest = $Dest }
  } else {
    Write-Host ("Junction OK -> {0}" -f ($link.Target -join ';'))
  }

  try {
    Remove-Item -LiteralPath $backup -Recurse -Force -ErrorAction Stop
    Write-Host 'Backup removed'
  } catch {
    Write-Host ("Backup remove partial: {0}" -f $_.Exception.Message)
    Get-ChildItem -LiteralPath $backup -Force -Recurse -ErrorAction SilentlyContinue |
      Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
  }

  return [PSCustomObject]@{ Label = $Label; Status = 'migrated'; GB = $size; Source = $Source; Dest = $Dest }
}

function Set-DevCacheEnv {
  param([string]$TargetDrive, [string]$UserName)
  $dev = "$TargetDrive`:\DevCache"
  $temp = "$TargetDrive`:\Temp\$UserName"
  $android = "$TargetDrive`:\Android"

  New-Item -ItemType Directory -Force -Path $temp, "$dev\npm-cache", "$dev\pip\Cache" | Out-Null

  [Environment]::SetEnvironmentVariable('GRADLE_USER_HOME', "$dev\gradle", 'User')
  [Environment]::SetEnvironmentVariable('PUB_CACHE', "$dev\pub\Cache", 'User')
  [Environment]::SetEnvironmentVariable('npm_config_cache', "$dev\npm-cache", 'User')
  [Environment]::SetEnvironmentVariable('PIP_CACHE_DIR', "$dev\pip\Cache", 'User')
  [Environment]::SetEnvironmentVariable('TEMP', $temp, 'User')
  [Environment]::SetEnvironmentVariable('TMP', $temp, 'User')
  [Environment]::SetEnvironmentVariable('ANDROID_HOME', $android, 'User')
  [Environment]::SetEnvironmentVariable('ANDROID_SDK_ROOT', $android, 'User')

  try {
    npm config set cache "$dev\npm-cache" --location=user 2>$null
    npm config set prefix (Join-Path $env:APPDATA 'npm') --location=user 2>$null
    Write-Host 'npm config updated'
  } catch {
    Write-Host 'npm config skipped'
  }

  $note = @"
# DevCache migration note
# Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')
# Junctions point C: paths to $dev so tools keep working.
# User env:
#   GRADLE_USER_HOME=$dev\gradle
#   PUB_CACHE=$dev\pub\Cache
#   npm_config_cache=$dev\npm-cache
#   PIP_CACHE_DIR=$dev\pip\Cache
#   TEMP/TMP=$temp
#   ANDROID_HOME/ANDROID_SDK_ROOT=$android
# Re-open terminals/IDEs for env to apply.
"@
  Set-Content -Path (Join-Path $dev 'README-migration.txt') -Value $note -Encoding UTF8
  Write-Host "Wrote $dev\README-migration.txt"
}

$beforeC = (Get-PSDrive C).Free
$beforeT = (Get-PSDrive $TargetDrive -ErrorAction SilentlyContinue).Free
Write-Host ("C free before: {0:N2} GB" -f ($beforeC / 1GB))
if ($null -ne $beforeT) { Write-Host ("{0}: free before: {1:N2} GB" -f $TargetDrive, ($beforeT / 1GB)) }

$results = @()

if ($UseDefaultDevCache) {
  $userProfile = Join-Path 'C:\Users' $UserName
  if (-not (Test-Path $userProfile)) { $userProfile = $env:USERPROFILE }
  $local = Join-Path $userProfile 'AppData\Local'
  $roaming = Join-Path $userProfile 'AppData\Roaming'
  $dev = "$TargetDrive`:\DevCache"

  if ($Execute) { New-Item -ItemType Directory -Force -Path $dev | Out-Null }

  $jobs = @(
    @{ Label = 'gradle'; Source = (Join-Path $userProfile '.gradle'); Dest = "$dev\gradle" }
    @{ Label = 'jdks'; Source = (Join-Path $userProfile '.jdks'); Dest = "$dev\jdks" }
    @{ Label = 'm2'; Source = (Join-Path $userProfile '.m2'); Dest = "$dev\m2" }
    @{ Label = 'miniconda3'; Source = (Join-Path $userProfile 'miniconda3'); Dest = "$dev\miniconda3" }
    @{ Label = 'npm-global'; Source = (Join-Path $roaming 'npm'); Dest = "$dev\npm-global" }
    @{ Label = 'npm-cache'; Source = (Join-Path $local 'npm-cache'); Dest = "$dev\npm-cache" }
    @{ Label = 'pub'; Source = (Join-Path $local 'Pub'); Dest = "$dev\pub" }
    @{ Label = 'develop'; Source = (Join-Path $userProfile 'develop'); Dest = "$dev\develop" }
    @{ Label = 'pip'; Source = (Join-Path $local 'pip'); Dest = "$dev\pip" }
    @{ Label = 'docker'; Source = (Join-Path $local 'Docker'); Dest = "$TargetDrive`:\Docker" }
    @{ Label = 'android'; Source = (Join-Path $local 'Android'); Dest = "$TargetDrive`:\Android" }
  )

  foreach ($j in $jobs) {
    $results += Migrate-ToJunction -Source $j.Source -Dest $j.Dest -Label $j.Label
  }
  if ($Execute) { Set-DevCacheEnv -TargetDrive $TargetDrive -UserName $UserName }
}
elseif ($Source -and $Dest) {
  $results += Migrate-ToJunction -Source $Source -Dest $Dest -Label $Label
  if ($Execute -and $ConfigureEnv) { Set-DevCacheEnv -TargetDrive $TargetDrive -UserName $UserName }
}
else {
  Write-Error 'Provide -Source/-Dest or -UseDefaultDevCache'
  exit 1
}

Write-Host "`n=== Migration results ===" -ForegroundColor Cyan
$results | Format-Table -AutoSize

$afterC = (Get-PSDrive C).Free
$afterT = (Get-PSDrive $TargetDrive -ErrorAction SilentlyContinue).Free
Write-Host ("C free after: {0:N2} GB (delta {1:N2} GB)" -f ($afterC / 1GB), (($afterC - $beforeC) / 1GB))
if ($null -ne $afterT -and $null -ne $beforeT) {
  Write-Host ("{0}: free after: {1:N2} GB (delta {2:N2} GB)" -f $TargetDrive, ($afterT / 1GB), (($afterT - $beforeT) / 1GB))
}
