@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0.."

echo ============================================================
echo  1/2  dotnet publish  (self-contained, win-x64)
echo ============================================================
dotnet publish NssfPayroll.csproj -c Release -r win-x64 --self-contained true -p:PublishReadyToRun=false -o bin\publish\win-x64
if errorlevel 1 (
    echo.
    echo Publish failed. Make sure the .NET 8 SDK is installed:
    echo https://dotnet.microsoft.com/download/dotnet/8.0
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  2/2  Inno Setup compile
echo ============================================================
set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" (
    echo Inno Setup 6 was not found at the default install path.
    echo Install it from https://jrsoftware.org/isdl.php , or edit the ISCC path in this file.
    pause
    exit /b 1
)

"%ISCC%" "Installer\NssfPayroll.iss"
if errorlevel 1 (
    echo.
    echo Inno Setup build failed - see the messages above.
    pause
    exit /b 1
)

echo.
echo Done: Installer\Output\NssfPayrollSetup.exe
pause
