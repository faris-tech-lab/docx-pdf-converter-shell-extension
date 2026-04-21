$ErrorActionPreference = "Stop"

$installDir = Join-Path $env:LOCALAPPDATA "DocxToPdf"
$converterSource = Join-Path $PSScriptRoot "convert.ps1"
$converterDest = Join-Path $installDir "convert.ps1"

# 1. Create install directory
if (-not (Test-Path $installDir)) {
    New-Item -ItemType Directory -Path $installDir -Force | Out-Null
}

# 2. Copy converter script and silent launcher
Copy-Item -Path $converterSource -Destination $converterDest -Force
$launcherSource = Join-Path $PSScriptRoot "launch.vbs"
$launcherDest = Join-Path $installDir "launch.vbs"
Copy-Item -Path $launcherSource -Destination $launcherDest -Force

# 3. Create registry entries
$regPath = "HKCU:\Software\Classes\SystemFileAssociations\.docx\shell\ConvertToPDF"
$cmdPath = "$regPath\command"

# Create the shell key
New-Item -Path $regPath -Force | Out-Null
Set-ItemProperty -Path $regPath -Name "(Default)" -Value "Convert to PDF"
Set-ItemProperty -Path $regPath -Name "Icon" -Value "shell32.dll,201"

# Create the command subkey
New-Item -Path $cmdPath -Force | Out-Null
$launcherPath = Join-Path $installDir "launch.vbs"
$command = "wscript.exe `"$launcherPath`" `"%1`""
Set-ItemProperty -Path $cmdPath -Name "(Default)" -Value $command

Write-Host "Installed successfully."
Write-Host "  Converter: $converterDest"
Write-Host "  Registry:  $regPath"
Write-Host "Right-click any .docx file to see 'Convert to PDF'."
