# build_installer.ps1
# This script publishes the WPF application in Release mode and compiles the setup installer.

$ProjectDir = $PSScriptRoot
if ($ProjectDir -eq "" -or $ProjectDir -eq $null) {
    $ProjectDir = Get-Location
}
$PublishDir = "$ProjectDir\bin\Release\net10.0-windows\publish"

Write-Host "1. Building and publishing the WPF application..." -ForegroundColor Green
dotnet publish "$ProjectDir\AccountingOcrTest.csproj" -c Release

if ($LASTEXITCODE -ne 0) {
    Write-Error "Dotnet publish failed!"
    exit 1
}

Write-Host "2. Copying default settings files (if present)..." -ForegroundColor Green
if (Test-Path "$ProjectDir\shippers.json") {
    Copy-Item -Path "$ProjectDir\shippers.json" -Destination "$PublishDir\" -Force
    Write-Host "   Copied shippers.json" -ForegroundColor Gray
}
if (Test-Path "$ProjectDir\excel_configs.json") {
    Copy-Item -Path "$ProjectDir\excel_configs.json" -Destination "$PublishDir\" -Force
    Write-Host "   Copied excel_configs.json" -ForegroundColor Gray
}

Write-Host "3. Packaging using Inno Setup Compiler..." -ForegroundColor Green
# Look for Inno Setup compiler in common directories
$ISCC = ""
$PotentialPaths = @(
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
    "C:\Program Files\Inno Setup 6\ISCC.exe",
    "C:\Program Files (x86)\Inno Setup 5\ISCC.exe",
    "C:\Program Files\Inno Setup 5\ISCC.exe"
)

foreach ($path in $PotentialPaths) {
    if (Test-Path $path) {
        $ISCC = $path
        break
    }
}

if ($ISCC -eq "") {
    # Try searching path
    $ISCC = Get-Command iscc -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source
}

if ($ISCC -eq "" -or $ISCC -eq $null) {
    Write-Host "`nWARNING: Inno Setup Compiler (ISCC.exe) was not found on your system." -ForegroundColor Yellow
    Write-Host "To compile the final installer, please:" -ForegroundColor Yellow
    Write-Host "  1. Download and install Inno Setup (free) from: https://jrsoftware.org/isinfo.php" -ForegroundColor Yellow
    Write-Host "  2. Open Inno Setup and compile the script file: '$ProjectDir\setup.iss'" -ForegroundColor Yellow
    Write-Host "`nThe application files have been successfully published and are ready to package in:" -ForegroundColor Green
    Write-Host "  $PublishDir" -ForegroundColor Green
    exit 0
}

Write-Host "Found Inno Setup Compiler at: $ISCC" -ForegroundColor Cyan
Write-Host "Compiling setup.iss..." -ForegroundColor Cyan

& $ISCC "$ProjectDir\setup.iss"

if ($LASTEXITCODE -eq 0) {
    Write-Host "`nSUCCESS! The installer has been created in:" -ForegroundColor Green
    Write-Host "  $ProjectDir\..\InstallerOutput\HFinSetup.exe" -ForegroundColor Green
} else {
    Write-Error "Inno Setup compilation failed!"
    exit 1
}
