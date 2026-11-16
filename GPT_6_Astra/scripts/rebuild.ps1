param()
$ErrorActionPreference = 'Stop'
$packageRoot = Split-Path -Parent $PSScriptRoot
$runtimeRoot = Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies'
$nodeExecutable = Join-Path $runtimeRoot 'node\bin\node.exe'
$nodePackages = Join-Path $runtimeRoot 'node\node_modules'
$moduleJunction = Join-Path $PSScriptRoot 'node_modules'
if (-not (Test-Path -LiteralPath $nodeExecutable)) { throw "Bundled Node runtime not found at $nodeExecutable" }
if (-not (Test-Path -LiteralPath $nodePackages)) { throw "Bundled Node packages not found at $nodePackages" }
if (-not (Test-Path -LiteralPath $moduleJunction)) {
    New-Item -ItemType Junction -Path $moduleJunction -Target $nodePackages | Out-Null
}
& $nodeExecutable (Join-Path $PSScriptRoot 'build_workbook.mjs')
if ($LASTEXITCODE -ne 0) { throw 'Workbook rebuild failed' }
Write-Output "Rebuilt workbook saved under $packageRoot\rebuild. Original deliverables were not overwritten."
