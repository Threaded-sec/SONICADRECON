# SonicADRecon Build Script
# PowerShell version

Write-Host "Building SonicADRecon..." -ForegroundColor Green

# Check if .NET 6.0 SDK is installed
try {
    $dotnetVersion = dotnet --version
    Write-Host "Found .NET SDK version: $dotnetVersion" -ForegroundColor Yellow
} catch {
    Write-Host "Error: .NET 6.0 SDK is not installed or not in PATH" -ForegroundColor Red
    Write-Host "Please install .NET 6.0 SDK from https://dotnet.microsoft.com/download" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

# Clean previous builds
Write-Host "Cleaning previous builds..." -ForegroundColor Yellow
if (Test-Path "ADAUDIT\bin") {
    Remove-Item "ADAUDIT\bin" -Recurse -Force
}
if (Test-Path "ADAUDIT\obj") {
    Remove-Item "ADAUDIT\obj" -Recurse -Force
}

# Restore packages
Write-Host "Restoring NuGet packages..." -ForegroundColor Yellow
dotnet restore ADAUDIT.sln

if ($LASTEXITCODE -ne 0) {
    Write-Host "Package restore failed!" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Build the solution
Write-Host "Building solution..." -ForegroundColor Yellow
dotnet build ADAUDIT.sln --configuration Release --no-restore

if ($LASTEXITCODE -ne 0) {
    Write-Host "Build failed!" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Check if executable was created
$exePath = "ADAUDIT\bin\Release\net6.0-windows\ADAUDIT.exe"
if (Test-Path $exePath) {
    $fileSize = (Get-Item $exePath).Length
    $fileSizeKB = [math]::Round($fileSize / 1KB, 2)
    
    Write-Host "Build completed successfully!" -ForegroundColor Green
    Write-Host "Executable location: $exePath" -ForegroundColor Cyan
    Write-Host "File size: $fileSizeKB KB" -ForegroundColor Cyan
    
    # Ask if user wants to run the application
    $runApp = Read-Host "Do you want to run the application now? (y/n)"
    if ($runApp -eq "y" -or $runApp -eq "Y") {
        Write-Host "Starting SonicADRecon..." -ForegroundColor Green
        Start-Process $exePath
    }
} else {
    Write-Host "Build completed but executable not found at expected location!" -ForegroundColor Red
}

Write-Host ""
Read-Host "Press Enter to exit" 