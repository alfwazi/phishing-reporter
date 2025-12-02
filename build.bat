@echo off
REM Build script for Phishing Reporter Outlook Add-in
REM Requires: Visual Studio 2017 or later with VSTO tools

echo ========================================
echo Phishing Reporter - Build Script
echo ========================================
echo.

REM Find MSBuild
set MSBUILD_PATH=
if exist "C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe" (
    set MSBUILD_PATH=C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe
) else if exist "C:\Program Files (x86)\Microsoft Visual Studio\2017\Professional\MSBuild\15.0\Bin\MSBuild.exe" (
    set MSBUILD_PATH=C:\Program Files (x86)\Microsoft Visual Studio\2017\Professional\MSBuild\15.0\Bin\MSBuild.exe
) else if exist "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe" (
    set MSBUILD_PATH=C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe
) else (
    echo ERROR: MSBuild not found. Please install Visual Studio.
    pause
    exit /b 1
)

echo Using MSBuild: %MSBUILD_PATH%
echo.

REM Clean previous builds
echo [1/3] Cleaning previous builds...
"%MSBUILD_PATH%" PhishingReporter.sln /t:Clean /p:Configuration=Release /p:Platform="Any CPU" /v:minimal
if errorlevel 1 (
    echo ERROR: Clean failed
    pause
    exit /b 1
)

REM Build the solution
echo [2/4] Building solution (Release mode)...
"%MSBUILD_PATH%" PhishingReporter\PhishingReporter.csproj /t:Build /p:Configuration=Release /p:Platform="Any CPU" /v:minimal /m
if errorlevel 1 (
    echo ERROR: Build failed
    pause
    exit /b 1
)

REM Build the installer
echo [3/4] Building installer...
if exist "C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\IDE\devenv.com" (
    "C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\IDE\devenv.com" PhishingReporter.sln /Build Release
) else if exist "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\Common7\IDE\devenv.com" (
    "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\Common7\IDE\devenv.com" PhishingReporter.sln /Build Release
) else if exist "C:\Program Files (x86)\Microsoft Visual Studio\2017\Community\Common7\IDE\devenv.com" (
    "C:\Program Files (x86)\Microsoft Visual Studio\2017\Community\Common7\IDE\devenv.com" PhishingReporter.sln /Build Release
) else (
    echo WARNING: Visual Studio IDE not found. Installer requires Visual Studio.
    echo Building DLL only. Installer can be built manually.
)

REM Create SCCM deployment package
echo [4/4] Creating SCCM deployment package...
if not exist "SCCM_Deployment" mkdir "SCCM_Deployment"
if not exist "SCCM_Deployment\Files" mkdir "SCCM_Deployment\Files"

xcopy /Y /I "PhishingReporter\bin\Release\*" "SCCM_Deployment\Files\"
if exist "Installer\Release\*.exe" (
    xcopy /Y "Installer\Release\*.exe" "SCCM_Deployment\"
)

REM Check if build succeeded
if exist "PhishingReporter\bin\Release\PhishingReporter.dll" (
    echo.
    echo ========================================
    echo BUILD SUCCESSFUL!
    echo ========================================
    echo.
    echo Output files:
    echo   - DLL: PhishingReporter\bin\Release\PhishingReporter.dll
    if exist "Installer\Release\*.exe" (
        echo   - Installer: Installer\Release\*.exe
    )
    echo   - SCCM Package: SCCM_Deployment\
    echo.
    echo Ready for enterprise deployment!
    echo.
) else (
    echo.
    echo ERROR: Build output not found
    pause
    exit /b 1
)

pause

