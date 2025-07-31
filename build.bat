@echo off
echo Building SonicADRecon...

REM Check if .NET 6.0 SDK is installed
dotnet --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: .NET 6.0 SDK is not installed or not in PATH
    echo Please install .NET 6.0 SDK from https://dotnet.microsoft.com/download
    pause
    exit /b 1
)

REM Clean previous builds
echo Cleaning previous builds...
if exist "ADAUDIT\bin" rmdir /s /q "ADAUDIT\bin"
if exist "ADAUDIT\obj" rmdir /s /q "ADAUDIT\obj"

REM Restore packages
echo Restoring NuGet packages...
dotnet restore ADAUDIT.sln

REM Build the solution
echo Building solution...
dotnet build ADAUDIT.sln --configuration Release --no-restore

if %errorlevel% neq 0 (
    echo Build failed!
    pause
    exit /b 1
)

echo Build completed successfully!
echo.
echo The executable is located at: ADAUDIT\bin\Release\net6.0-windows\ADAUDIT.exe
echo.
pause 