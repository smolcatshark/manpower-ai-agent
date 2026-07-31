@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo.
echo ================================================
echo  Manpower Report Agent - Windows Build
echo ================================================
echo.

where py >nul 2>nul
if errorlevel 1 (
    echo Python Launcher was not found.
    echo Install Python 3.12 and try again.
    pause
    exit /b 1
)

if not exist ".venv-desktop\Scripts\python.exe" (
    echo Creating desktop build environment...
    py -3.12 -m venv .venv-desktop
    if errorlevel 1 goto :build_failed
)

call ".venv-desktop\Scripts\activate.bat"
if errorlevel 1 goto :build_failed

echo Installing build dependencies...
python -m pip install --upgrade pip
if errorlevel 1 goto :build_failed

python -m pip install -r requirements-desktop.txt
if errorlevel 1 goto :build_failed

echo Removing old build output...
if exist build rmdir /s /q build
if exist "dist\Manpower Report Agent" (
    rmdir /s /q "dist\Manpower Report Agent"
)

echo Building the desktop application...
python -m PyInstaller ^
    --noconfirm ^
    --clean ^
    Manpower_Report_Agent.spec

if errorlevel 1 goto :build_failed

echo.
echo Build completed successfully.
echo App location:
echo "%CD%\dist\Manpower Report Agent\Manpower Report Agent.exe"
echo.
pause
exit /b 0

:build_failed
echo.
echo Build failed. Review the messages above.
echo.
pause
exit /b 1
