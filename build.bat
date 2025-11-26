@echo off
chcp 65001 > nul

echo ========================================
echo      Grade Analysis System - Builder
echo ========================================
echo.

pip show pyinstaller > nul 2>&1
if errorlevel 1 (
    echo Installing PyInstaller...
    pip install pyinstaller
)

echo.
echo Building...
echo.

if exist "dist" rd /s /q dist
if exist "build" rd /s /q build

pyinstaller --noupx build.spec

if errorlevel 1 (
    echo.
    echo Build failed!
    pause
    exit /b 1
)

echo.
echo ========================================
echo Build complete!
echo Output: dist\成绩分析系统.exe
echo ========================================
echo.

pause

