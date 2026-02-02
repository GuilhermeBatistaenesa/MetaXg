param(
    [string]$SpecPath = "build\\metaxg_onedir.spec"
)

$ErrorActionPreference = "Stop"

$python = "python"
if (Test-Path ".venv\\Scripts\\python.exe") {
    $python = ".venv\\Scripts\\python.exe"
}

if (-not (Test-Path $SpecPath)) {
    throw "Spec nao encontrado: $SpecPath"
}

if (Test-Path "dist\\MetaXg") {
    Remove-Item -Recurse -Force "dist\\MetaXg"
}

Write-Host "[build_windows] executando PyInstaller..."
& $python -m PyInstaller $SpecPath --clean

Write-Host "[build_windows] concluido. Saida em dist\\MetaXg\\"

if (Test-Path ".env") {
    if (Test-Path "dist\\MetaXg") {
        Copy-Item ".env" "dist\\MetaXg\\.env" -Force
    }
    if (Test-Path "dist\\MetaXg.exe") {
        Copy-Item ".env" "dist\\.env" -Force
    }
    Write-Host "[build_windows] .env copiado para dist."
}
