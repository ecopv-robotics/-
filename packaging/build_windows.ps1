$ErrorActionPreference = 'Stop'

$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$releaseName = 'ECOPV_Mail_Audit_Workbench_2026-10-21'
$releaseDir = Join-Path (Join-Path $projectRoot 'dist') $releaseName
$archivePath = Join-Path (Join-Path $projectRoot 'dist') ($releaseName + '.zip')
$buildPath = Join-Path (Join-Path $projectRoot 'build') 'pyinstaller-20261021'
$ocrModelNames = @(
    'PP-LCNet_x1_0_doc_ori',
    'PP-LCNet_x1_0_textline_ori',
    'PP-OCRv6_medium_det',
    'PP-OCRv6_medium_rec',
    'UVDoc'
)
$ocrModelSource = Join-Path $env:USERPROFILE '.paddlex\official_models'
$sevenZipCandidates = @(
    (Join-Path $env:ProgramFiles '7-Zip\7z.exe'),
    (Join-Path ${env:ProgramFiles(x86)} '7-Zip\7z.exe')
)
$sevenZipPath = $sevenZipCandidates | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
if (-not $sevenZipPath) {
    throw '7z.exe is required for safe RAR extraction; install 7-Zip before building this release.'
}
$sevenZipDir = Split-Path -Parent $sevenZipPath
$sevenZipDll = Join-Path $sevenZipDir '7z.dll'
$sevenZipLicense = Join-Path $sevenZipDir 'License.txt'
foreach ($requiredFile in @($sevenZipDll, $sevenZipLicense)) {
    if (-not (Test-Path -LiteralPath $requiredFile)) {
        throw "Required 7-Zip runtime/license file is missing: $requiredFile"
    }
}
foreach ($modelName in $ocrModelNames) {
    $modelPath = Join-Path $ocrModelSource $modelName
    if (-not (Test-Path -LiteralPath $modelPath -PathType Container)) {
        throw "Required PaddleOCR model is missing: $modelPath"
    }
    if (-not (Get-ChildItem -LiteralPath $modelPath -Recurse -File | Select-Object -First 1)) {
        throw "Required PaddleOCR model folder is empty: $modelPath"
    }
}

if (Test-Path -LiteralPath $releaseDir) {
    throw "Release folder already exists; preserving it: $releaseDir"
}
if (Test-Path -LiteralPath $archivePath) {
    throw "Release archive already exists; preserving it: $archivePath"
}

Push-Location $projectRoot
$previousPythonPath = $env:PYTHONPATH
$previousPackagingFlag = $env:ECOPV_PACKAGING_BUILD
try {
    $env:PYTHONPATH = Join-Path $PSScriptRoot 'build_support'
    $env:ECOPV_PACKAGING_BUILD = '1'
    & python -m PyInstaller --noconfirm --workpath $buildPath --distpath (Join-Path $projectRoot 'dist') (Join-Path $PSScriptRoot 'mail_audit_workbench.spec')
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller failed with exit code $LASTEXITCODE"
    }
} finally {
    Pop-Location
    $env:PYTHONPATH = $previousPythonPath
    $env:ECOPV_PACKAGING_BUILD = $previousPackagingFlag
}

if (-not (Test-Path -LiteralPath $releaseDir)) {
    throw "Expected PyInstaller output was not created: $releaseDir"
}

Copy-Item -LiteralPath (Join-Path $projectRoot 'config.example.yaml') -Destination (Join-Path $releaseDir 'config.example.yaml')
Copy-Item -LiteralPath (Join-Path $projectRoot 'config.example.yaml') -Destination (Join-Path $releaseDir 'config.yaml')
Copy-Item -LiteralPath (Join-Path $projectRoot 'workbench.html') -Destination (Join-Path $releaseDir 'workbench.html')
Copy-Item -LiteralPath (Join-Path $projectRoot 'web_console') -Destination $releaseDir -Recurse
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'start_workbench.bat') -Destination (Join-Path $releaseDir '启动工作台.bat')
Copy-Item -LiteralPath (Join-Path $PSScriptRoot '操作使用说明.md') -Destination (Join-Path $releaseDir '操作使用说明.md')
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'assets\业务测试反馈表.xlsx') -Destination (Join-Path $releaseDir '业务测试反馈表.xlsx')
$ocrModelTarget = Join-Path $releaseDir 'model_cache\official_models'
New-Item -ItemType Directory -Path $ocrModelTarget -Force | Out-Null
foreach ($modelName in $ocrModelNames) {
    Copy-Item -LiteralPath (Join-Path $ocrModelSource $modelName) -Destination $ocrModelTarget -Recurse
}
$toolsDir = Join-Path $releaseDir 'tools'
New-Item -ItemType Directory -Path $toolsDir | Out-Null
Copy-Item -LiteralPath $sevenZipPath -Destination $toolsDir
Copy-Item -LiteralPath $sevenZipDll -Destination $toolsDir
Copy-Item -LiteralPath $sevenZipLicense -Destination (Join-Path $toolsDir '7-Zip-License.txt')

Push-Location (Join-Path $projectRoot 'dist')
try {
    & $sevenZipPath a -tzip -mcu=on -mx=1 ($releaseName + '.zip') $releaseName
    if ($LASTEXITCODE -ne 0) {
        throw "7-Zip archive creation failed with exit code $LASTEXITCODE"
    }
} finally {
    Pop-Location
}
Write-Output "Release folder: $releaseDir"
Write-Output "Release archive: $archivePath"
