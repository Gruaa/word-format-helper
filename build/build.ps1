$ErrorActionPreference = "Stop"

$root = "d:\Trae\word plugin"
$msbuild = "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
$addinProj = Join-Path $root "src\WordFormatHelper.AddIn\WordFormatHelper.AddIn.csproj"
$installerProj = Join-Path $root "src\WordFormatHelper.Installer\WordFormatHelper.Installer.csproj"
$addinBin = Join-Path $root "src\WordFormatHelper.AddIn\bin\Release\WordFormatHelper.AddIn.dll"
$installerResDir = Join-Path $root "src\WordFormatHelper.Installer\Resources"
$installerResDll = Join-Path $installerResDir "WordFormatHelper.AddIn.dll"
$distDir = Join-Path $root "dist"
$distExe = Join-Path $distDir "WordFormatHelperSetup.exe"
$installerBin = Join-Path $root "src\WordFormatHelper.Installer\bin\Release\WordFormatHelperSetup.exe"

Write-Host "==> [1/4] Building Add-in DLL..." -ForegroundColor Cyan
& $msbuild $addinProj /p:Configuration=Release /v:minimal /m
if ($LASTEXITCODE -ne 0) { throw "Add-in build failed (exit $LASTEXITCODE)" }

if (-not (Test-Path $addinBin)) { throw "Build output not found: $addinBin" }
Write-Host "    Output: $addinBin" -ForegroundColor Green

Write-Host "==> [2/4] Copying DLL to installer resources..." -ForegroundColor Cyan
if (-not (Test-Path $installerResDir)) { New-Item -ItemType Directory -Path $installerResDir | Out-Null }
Copy-Item -Path $addinBin -Destination $installerResDll -Force
Write-Host "    Copied to: $installerResDll" -ForegroundColor Green

Write-Host "==> [3/4] Building installer exe..." -ForegroundColor Cyan
& $msbuild $installerProj /p:Configuration=Release /v:minimal /m
if ($LASTEXITCODE -ne 0) { throw "Installer build failed (exit $LASTEXITCODE)" }

if (-not (Test-Path $installerBin)) { throw "Installer output not found: $installerBin" }
Write-Host "    Output: $installerBin" -ForegroundColor Green

Write-Host "==> [4/4] Publishing to dist..." -ForegroundColor Cyan
if (-not (Test-Path $distDir)) { New-Item -ItemType Directory -Path $distDir | Out-Null }
Copy-Item -Path $installerBin -Destination $distExe -Force
Write-Host "    Final: $distExe" -ForegroundColor Green

Write-Host ""
Write-Host "BUILD COMPLETE" -ForegroundColor Yellow
Write-Host "  Setup: $distExe" -ForegroundColor Yellow
Write-Host "  Install: run as administrator" -ForegroundColor Yellow
Write-Host "  Uninstall: WordFormatHelperSetup.exe /uninstall" -ForegroundColor Yellow
