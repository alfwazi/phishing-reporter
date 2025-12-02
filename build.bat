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
echo [2/3] Building solution (Release mode)...
"%MSBUILD_PATH%" PhishingReporter.sln /t:Build /p:Configuration=Release /p:Platform="Any CPU" /v:minimal /m
if errorlevel 1 (
    echo ERROR: Build failed
    pause
    exit /b 1
)

REM Check if build succeeded
if exist "PhishingReporter\bin\Release\PhishingReporter.dll" (
    echo.
    echo ========================================
    echo BUILD SUCCESSFUL!
    echo ========================================
    echo.
    echo Output: PhishingReporter\bin\Release\PhishingReporter.dll
    echo.
    echo Next steps:
    echo 1. Build the Installer project in Visual Studio
    echo 2. Install using the generated installer from Installer\Release\
    echo.
) else (
    echo.
    echo ERROR: Build output not found
    pause
    exit /b 1
)

pause

