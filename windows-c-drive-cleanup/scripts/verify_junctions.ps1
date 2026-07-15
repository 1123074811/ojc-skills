<#
.SYNOPSIS
  Verify junctions, env vars, free space, and basic tool smoke checks.
.EXAMPLE
  pwsh -NoProfile -File .\verify_junctions.ps1 -TargetDrive D -UserName MECHREVO
#>
[CmdletBinding()]
param(
  [string]$TargetDrive = 'D',
  [string]$UserName = $env:USERNAME
)

$ErrorActionPreference = 'SilentlyContinue'

function Get-DirSizeGB([string]$Path) {
  if (-not (Test-Path -LiteralPath $Path)) { return $null }
  $sum = (Get-ChildItem -LiteralPath $Path -Force -Recurse -File -ErrorAction SilentlyContinue |
    Measure-Object Length -Sum).Sum
  if ($null -eq $sum) { return 0 }
  return [math]::Round($sum / 1GB, 2)
}

Write-Host '=== Free space ===' -ForegroundColor Cyan
Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Name -match '^[C-Z]$' } | ForEach-Object {
  [PSCustomObject]@{
    Name   = $_.Name
    FreeGB = [math]::Round($_.Free / 1GB, 2)
    UsedGB = [math]::Round($_.Used / 1GB, 2)
  }
} | Format-Table -AutoSize

$userProfile = Join-Path 'C:\Users' $UserName
if (-not (Test-Path $userProfile)) { $userProfile = $env:USERPROFILE }
$local = Join-Path $userProfile 'AppData\Local'
$roaming = Join-Path $userProfile 'AppData\Roaming'
$dev = "$TargetDrive`:\DevCache"

$paths = @(
  @{ N = 'Docker'; P = (Join-Path $local 'Docker'); Expect = "$TargetDrive`:\Docker" }
  @{ N = 'Android'; P = (Join-Path $local 'Android'); Expect = "$TargetDrive`:\Android" }
  @{ N = 'gradle'; P = (Join-Path $userProfile '.gradle'); Expect = "$dev\gradle" }
  @{ N = 'jdks'; P = (Join-Path $userProfile '.jdks'); Expect = "$dev\jdks" }
  @{ N = 'm2'; P = (Join-Path $userProfile '.m2'); Expect = "$dev\m2" }
  @{ N = 'miniconda3'; P = (Join-Path $userProfile 'miniconda3'); Expect = "$dev\miniconda3" }
  @{ N = 'npm-global'; P = (Join-Path $roaming 'npm'); Expect = "$dev\npm-global" }
  @{ N = 'npm-cache'; P = (Join-Path $local 'npm-cache'); Expect = "$dev\npm-cache" }
  @{ N = 'pub'; P = (Join-Path $local 'Pub'); Expect = "$dev\pub" }
  @{ N = 'develop'; P = (Join-Path $userProfile 'develop'); Expect = "$dev\develop" }
  @{ N = 'pip'; P = (Join-Path $local 'pip'); Expect = "$dev\pip" }
)

Write-Host '=== Junction status ===' -ForegroundColor Cyan
foreach ($x in $paths) {
  if (-not (Test-Path -LiteralPath $x.P)) {
    '{0,-12} MISSING' -f $x.N
    continue
  }
  $i = Get-Item -LiteralPath $x.P -Force
  $sz = Get-DirSizeGB $x.P
  $target = ($i.Target -join ';')
  $ok = ($i.LinkType -eq 'Junction')
  $mark = if ($ok) { 'OK' } else { 'NOT-LINK' }
  '{0,-12} {1,6} {2,7}GB  Link={3,-10} -> {4}' -f $x.N, $mark, $sz, $i.LinkType, $target
}

Write-Host "`n=== Leftover *.pre-migrate-bak ===" -ForegroundColor Cyan
$baks = @()
$baks += Get-ChildItem -LiteralPath $userProfile -Force -Directory -Filter '*.pre-migrate-bak' -ErrorAction SilentlyContinue
$baks += Get-ChildItem -LiteralPath $local -Force -Directory -Filter '*.pre-migrate-bak' -ErrorAction SilentlyContinue
$baks += Get-ChildItem -LiteralPath $roaming -Force -Directory -Filter '*.pre-migrate-bak' -ErrorAction SilentlyContinue
if ($baks) { $baks | Select-Object FullName | Format-Table -AutoSize } else { Write-Host 'None' }

Write-Host "`n=== User environment variables ===" -ForegroundColor Cyan
foreach ($n in @(
    'GRADLE_USER_HOME', 'PUB_CACHE', 'npm_config_cache', 'PIP_CACHE_DIR',
    'TEMP', 'TMP', 'ANDROID_HOME', 'ANDROID_SDK_ROOT'
  )) {
  $v = [Environment]::GetEnvironmentVariable($n, 'User')
  '{0}={1}' -f $n, $v
}

Write-Host "`n=== Smoke checks ===" -ForegroundColor Cyan
if (Get-Command claude -ErrorAction SilentlyContinue) {
  Write-Host ('claude: ' + (claude --version 2>$null))
} else { Write-Host 'claude: not in PATH' }

if (Get-Command npm -ErrorAction SilentlyContinue) {
  Write-Host ('npm prefix: ' + (npm config get prefix 2>$null))
  Write-Host ('npm cache : ' + (npm config get cache 2>$null))
  Write-Host ('npm root-g: ' + (npm root -g 2>$null))
} else { Write-Host 'npm: not in PATH' }

if (Get-Command java -ErrorAction SilentlyContinue) {
  $javaOut = & java -version 2>&1 | Select-Object -First 1
  Write-Host ("java: $javaOut")
} else { Write-Host 'java: not in PATH' }

$flutter = Join-Path $userProfile 'develop\flutter\bin\flutter.bat'
Write-Host ('flutter.bat exists: ' + (Test-Path -LiteralPath $flutter))

Write-Host "`nVerify complete." -ForegroundColor Green
