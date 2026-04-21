$ErrorActionPreference = "Stop"

$installDir = Join-Path $env:LOCALAPPDATA "DocxToPdf"
$regPath = "HKCU:\Software\Classes\SystemFileAssociations\.docx\shell\ConvertToPDF"

# 1. Remove registry key
if (Test-Path $regPath) {
    Remove-Item -Path $regPath -Recurse -Force
    Write-Host "Registry key removed: $regPath"
} else {
    Write-Host "Registry key not found (already removed)."
}

# 2. Remove install directory
if (Test-Path $installDir) {
    Remove-Item -Path $installDir -Recurse -Force
    Write-Host "Install directory removed: $installDir"
} else {
    Write-Host "Install directory not found (already removed)."
}

Write-Host "Uninstall complete. Context menu entry is gone."
