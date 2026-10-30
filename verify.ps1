# GATE leve do Dashboard HTML: node --check + py_compile -> exit 0/1
$ErrorActionPreference = "Stop"
$root = $PSScriptRoot

$jsFiles = @(
    "assets\app.js",
    "assets\excel.js",
    "assets\regras.js",
    "assets\snapshot-template.js"
)

Write-Host "==> JS (node --check)" -ForegroundColor Cyan
foreach ($f in $jsFiles) {
    $p = Join-Path $root $f
    if (-not (Test-Path -LiteralPath $p)) { Write-Host "FALTANDO: $f" -ForegroundColor Red; exit 1 }
    node --check $p 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) { Write-Host "FALHOU: $f" -ForegroundColor Red; exit 1 }
}

Write-Host "==> Python (py_compile)" -ForegroundColor Cyan
python -m py_compile (Join-Path $root "build_snapshot_template.py")
if ($LASTEXITCODE -ne 0) { Write-Host "FALHOU: py_compile" -ForegroundColor Red; exit 1 }

Write-Host "OK" -ForegroundColor Green
exit 0
