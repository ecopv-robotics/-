<#
.SYNOPSIS
Prepare a PyInstaller release directory for external delivery (strip credentials and runtime residue).

.DESCRIPTION
Problem it solves: smoke tests run directly inside the release directory leave a real
config.yaml (plain-text mailbox / workorder passwords) plus logs/, output/, storage/
inside the package, while the operator manual states "this package contains no real
account configuration".

Rules:
  * Move, never delete - residue goes to _residue_backup_<timestamp> OUTSIDE the release dir.
  * Hard assertion that the backup dir is not inside the release dir.
  * Refuse delivery when config.yaml still holds non-empty credentials, unless -ResetConfig.

.EXAMPLE
powershell -ExecutionPolicy Bypass -File packaging\prepare_delivery.ps1
powershell -ExecutionPolicy Bypass -File packaging\prepare_delivery.ps1 -ResetConfig
#>
[CmdletBinding()]
param(
    [string]$ReleaseDir = "",
    [switch]$ResetConfig
)

$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($ReleaseDir)) {
    $ReleaseDir = Join-Path (Join-Path $projectRoot 'dist') 'ECOPV_Mail_Audit_Workbench_2026-10-10-r2'
}
$ReleaseDir = (Resolve-Path -LiteralPath $ReleaseDir).Path

$distRoot = Split-Path -Parent $ReleaseDir
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$backupDir = Join-Path $distRoot ("_residue_backup_" + $stamp)

# Defensive: backup must never land inside the release dir.
if ($backupDir.StartsWith($ReleaseDir, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Backup dir must not be inside the release dir: $backupDir"
}

Write-Host "release : $ReleaseDir"
Write-Host "backup  : $backupDir"
Write-Host ""

# ---- 1) credential check ----
$configPath = Join-Path $ReleaseDir 'config.yaml'
$dirty = @()
if (Test-Path -LiteralPath $configPath) {
    foreach ($line in Get-Content -LiteralPath $configPath) {
        if ($line -match '^\s*(password|address|username)\s*:\s*(.+?)\s*$') {
            $key = $Matches[1]
            $val = $Matches[2].Trim().Trim("'").Trim('"')
            if ($val -and $val -ne '[]' -and $val -ne '{}') {
                $dirty += $key
            }
        }
    }
}

if ($dirty.Count -gt 0) {
    if ($ResetConfig) {
        $example = Join-Path $ReleaseDir 'config.example.yaml'
        if (-not (Test-Path -LiteralPath $example)) {
            $example = Join-Path $projectRoot 'config.example.yaml'
        }
        Copy-Item -LiteralPath $example -Destination $configPath -Force
        Write-Host "[reset ] config.yaml overwritten with the blank template (was: $($dirty -join ', '))" -ForegroundColor Yellow
    } else {
        Write-Host "[block ] config.yaml still holds credentials: $($dirty -join ', ')" -ForegroundColor Red
        Write-Host "         Re-run with -ResetConfig to regenerate a blank config."
        exit 3
    }
} else {
    Write-Host "[  ok  ] config.yaml has no plain-text credentials" -ForegroundColor Green
}

# ---- 2) move runtime residue out ----
$moved = @()
foreach ($name in @('logs', 'output', 'storage', 'cache', 'startup-smoke.log', 'startup-smoke-error.log')) {
    $src = Join-Path $ReleaseDir $name
    if (Test-Path -LiteralPath $src) {
        if (-not (Test-Path -LiteralPath $backupDir)) {
            New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
        }
        Move-Item -LiteralPath $src -Destination (Join-Path $backupDir $name) -Force
        $moved += $name
    }
}
if ($moved.Count -gt 0) {
    Write-Host "[moved ] $($moved -join ', ')" -ForegroundColor Yellow
} else {
    Write-Host "[  ok  ] no runtime residue" -ForegroundColor Green
}

# ---- 3) report only: reference tables imported by the operator ----
$dataDir = Join-Path $ReleaseDir 'data'
if (Test-Path -LiteralPath $dataDir) {
    Write-Host "[ note ] data/ exists (operator-imported reference tables); not touched, confirm manually." -ForegroundColor Cyan
}

# ---- 4) final listing ----
Write-Host ""
Write-Host "=== release dir top level ==="
Get-ChildItem -LiteralPath $ReleaseDir | ForEach-Object { Write-Host ("  " + $_.Name) }
Write-Host ""
Write-Host "Done. Zip this directory for delivery; never ship _residue_backup_* folders."
