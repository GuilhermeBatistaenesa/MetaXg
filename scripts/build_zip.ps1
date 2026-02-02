param(
    [string]$Version = "",
    [ValidateSet("patch", "minor", "major")]
    [string]$Bump = "patch",
    [string]$AppName = "MetaXg",
    [string]$DistDir = "dist",
    [string]$ReleaseDir = "P:\\ProcessoMetaX\\releases"
)

$ErrorActionPreference = "Stop"

function Parse-Version([string]$v) {
    if ($v -notmatch "^\d+\.\d+\.\d+$") {
        throw "Versao invalida: $v"
    }
    $parts = $v.Split(".") | ForEach-Object { [int]$_ }
    return $parts
}

function Bump-Version([string]$v, [string]$bump) {
    $parts = Parse-Version $v
    $major = $parts[0]
    $minor = $parts[1]
    $patch = $parts[2]
    switch ($bump) {
        "major" { $major++; $minor = 0; $patch = 0 }
        "minor" { $minor++; $patch = 0 }
        default { $patch++ }
    }
    return "$major.$minor.$patch"
}

$versionFile = "version.txt"
$versionFile2 = "VERSION.txt"
$current = "0.0.0"
if (Test-Path $versionFile) {
    $current = (Get-Content $versionFile -ErrorAction SilentlyContinue | Select-Object -First 1).Trim()
} elseif (Test-Path $versionFile2) {
    $current = (Get-Content $versionFile2 -ErrorAction SilentlyContinue | Select-Object -First 1).Trim()
}

if (-not $Version) {
    $Version = Bump-Version $current $Bump
}

Write-Host "[build_zip] versao atual: $current"
Write-Host "[build_zip] versao nova: $Version"

Set-Content -Path $versionFile -Value $Version -Encoding UTF8
Set-Content -Path $versionFile2 -Value $Version -Encoding UTF8

$sourceDir = Join-Path $DistDir $AppName
if (-not (Test-Path $sourceDir)) {
    throw "Diretorio de build nao encontrado: $sourceDir"
}

if (-not (Test-Path $ReleaseDir)) {
    New-Item -Path $ReleaseDir -ItemType Directory | Out-Null
}

$zipName = "${AppName}_${Version}.zip"
$shaName = "${AppName}_${Version}.sha256"
$zipPath = Join-Path $ReleaseDir $zipName
$shaPath = Join-Path $ReleaseDir $shaName

if (Test-Path $zipPath) { Remove-Item $zipPath -Force }

Write-Host "[build_zip] gerando zip: $zipPath"
Compress-Archive -Path (Join-Path $sourceDir "*") -DestinationPath $zipPath

$hash = (Get-FileHash -Path $zipPath -Algorithm SHA256).Hash.ToLower()
Set-Content -Path $shaPath -Value "$hash  $zipName" -Encoding UTF8

$latest = @{
    app_name          = $AppName
    version           = $Version
    package_filename  = $zipName
    sha256_filename   = $shaName
    zip_name          = $zipName
    sha256_name       = $shaName
    published_at      = (Get-Date).ToString("s")
}

$latestPath = Join-Path $ReleaseDir "latest.json"
$latest | ConvertTo-Json -Depth 3 | Set-Content -Path $latestPath -Encoding UTF8

Write-Host "[build_zip] ok: $zipName"
Write-Host "[build_zip] ok: $shaName"
Write-Host "[build_zip] ok: latest.json"
