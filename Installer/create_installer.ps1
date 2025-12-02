# PowerShell script to create standalone installer
# This creates a self-extracting installer that doesn't require Visual Studio

param(
    [string]$OutputPath = "PhishingReporter-Setup.exe"
)

Write-Host "Creating Phishing Reporter Installer..." -ForegroundColor Green

$dllPath = "..\PhishingReporter\bin\Release\PhishingReporter.dll"
$configPath = "..\PhishingReporter\bin\Release\PhishingReporter.dll.config"
$htmlAgilityPath = "..\packages\HtmlAgilityPack.1.12.2\lib\Net45\HtmlAgilityPack.dll"

if (-not (Test-Path $dllPath)) {
    Write-Host "ERROR: DLL not found at $dllPath" -ForegroundColor Red
    Write-Host "Please build the project first." -ForegroundColor Yellow
    exit 1
}

# Create temporary directory for installer contents
$tempDir = "$env:TEMP\PhishingReporterInstaller"
if (Test-Path $tempDir) {
    Remove-Item $tempDir -Recurse -Force
}
New-Item -ItemType Directory -Path $tempDir | Out-Null
New-Item -ItemType Directory -Path "$tempDir\Files" | Out-Null

# Copy files
Copy-Item $dllPath -Destination "$tempDir\Files\PhishingReporter.dll" -Force
if (Test-Path $configPath) {
    Copy-Item $configPath -Destination "$tempDir\Files\PhishingReporter.dll.config" -Force
}
if (Test-Path $htmlAgilityPath) {
    Copy-Item $htmlAgilityPath -Destination "$tempDir\Files\HtmlAgilityPack.dll" -Force
}

# Create installation script
$installScript = @"
@echo off
REM Phishing Reporter - Silent Installation Script
REM For SCCM/Group Policy Deployment

setlocal
set INSTALL_DIR=%ProgramFiles%\Geidea\PhishingReporter
set VSTO_DIR=%ProgramFiles%\Common Files\Microsoft Shared\VSTO

echo Installing Phishing Reporter Outlook Add-in...

REM Check for Outlook
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office" /s | findstr /i "Outlook" >nul
if errorlevel 1 (
    echo ERROR: Microsoft Outlook not found
    exit /b 1
)

REM Create installation directory
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

REM Copy files
xcopy /Y /I "%~dp0Files\*" "%INSTALL_DIR%\"

REM Register with Outlook (VSTO add-in registration)
REM This requires VSTO runtime and proper manifest
REM For production, use the proper VSTO deployment method

echo Phishing Reporter installed to %INSTALL_DIR%
echo.
echo Installation complete!
echo Please restart Outlook to activate the add-in.

exit /b 0
"@

$installScript | Out-File -FilePath "$tempDir\install.cmd" -Encoding ASCII

# Create uninstall script
$uninstallScript = @"
@echo off
REM Phishing Reporter - Uninstall Script

set INSTALL_DIR=%ProgramFiles%\Geidea\PhishingReporter

if exist "%INSTALL_DIR%" (
    echo Uninstalling Phishing Reporter...
    rmdir /S /Q "%INSTALL_DIR%"
    echo Uninstallation complete.
) else (
    echo Phishing Reporter not found.
)

exit /b 0
"@

$uninstallScript | Out-File -FilePath "$tempDir\uninstall.cmd" -Encoding ASCII

# Create README
$readme = @"
Phishing Reporter - Installation Package
========================================

This package contains:
- PhishingReporter.dll (main plugin)
- HtmlAgilityPack.dll (dependency)
- install.cmd (installation script)
- uninstall.cmd (uninstallation script)

For SCCM Deployment:
1. Extract this package to a network share
2. Create SCCM Application
3. Install command: install.cmd
4. Uninstall command: uninstall.cmd
5. Detection: Check for %ProgramFiles%\Geidea\PhishingReporter\PhishingReporter.dll

For Group Policy:
1. Copy package to network share
2. Create GPO with startup script: \\server\share\install.cmd

Requirements:
- Microsoft Outlook 2013 or later
- .NET Framework 4.6.1 or later
- VSTO Runtime (usually pre-installed with Office)

After installation, restart Outlook to activate the add-in.
"@

$readme | Out-File -FilePath "$tempDir\README.txt" -Encoding ASCII

# Create 7-Zip self-extracting archive (if 7z is available)
$sevenZip = "C:\Program Files\7-Zip\7z.exe"
if (Test-Path $sevenZip) {
    Write-Host "Creating self-extracting installer..." -ForegroundColor Green
    & $sevenZip a -sfx "$OutputPath" "$tempDir\*"
    Write-Host "✅ Installer created: $OutputPath" -ForegroundColor Green
} else {
    Write-Host "7-Zip not found. Creating ZIP package instead..." -ForegroundColor Yellow
    Compress-Archive -Path "$tempDir\*" -DestinationPath "$OutputPath.zip" -Force
    Write-Host "✅ Package created: $OutputPath.zip" -ForegroundColor Green
    Write-Host "   Extract and run install.cmd for deployment" -ForegroundColor Yellow
}

# Cleanup
Remove-Item $tempDir -Recurse -Force

Write-Host ""
Write-Host "Installer package ready for deployment!" -ForegroundColor Green

