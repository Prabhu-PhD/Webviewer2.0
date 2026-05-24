[CmdletBinding()]
param(
    [string]$ShareName = "Webviewer2Catalog",
    [string]$CatalogDirectory = "catalog"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$catalogPath = Join-Path $projectRoot $CatalogDirectory
$manifestSource = Join-Path $projectRoot "manifest.localhost.xml"
$manifestTarget = Join-Path $catalogPath "manifest.localhost.xml"
$catalogId = "{8D9960B4-0B6B-4A9B-8B58-11AB6F2E8F73}"
$trustedCatalogKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$catalogId"
$uncPath = "\\localhost\$ShareName"

if (-not (Test-Path -LiteralPath $manifestSource)) {
    throw "manifest.localhost.xml not found at $manifestSource"
}

New-Item -ItemType Directory -Force -Path $catalogPath | Out-Null
Copy-Item -LiteralPath $manifestSource -Destination $manifestTarget -Force

$existingShare = Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue
if ($null -eq $existingShare) {
    New-SmbShare -Name $ShareName -Path $catalogPath -ReadAccess "Everyone" | Out-Null
} elseif ($existingShare.Path -ne $catalogPath) {
    throw "An SMB share named '$ShareName' already exists and points to '$($existingShare.Path)'."
}

New-Item -Path $trustedCatalogKey -Force | Out-Null
New-ItemProperty -Path $trustedCatalogKey -Name "Id" -PropertyType String -Value $catalogId -Force | Out-Null
New-ItemProperty -Path $trustedCatalogKey -Name "Url" -PropertyType String -Value $uncPath -Force | Out-Null
New-ItemProperty -Path $trustedCatalogKey -Name "Flags" -PropertyType DWord -Value 1 -Force | Out-Null

Write-Host "Shared-folder catalog is ready." -ForegroundColor Green
Write-Host "Local folder: $catalogPath"
Write-Host "UNC path: $uncPath"
Write-Host ""
Write-Host "Next in PowerPoint desktop:"
Write-Host "  1. Restart PowerPoint."
Write-Host "  2. Home > Add-ins > Advanced."
Write-Host "  3. SHARED FOLDER."
Write-Host "  4. Add 'Slide Web Viewer'."
