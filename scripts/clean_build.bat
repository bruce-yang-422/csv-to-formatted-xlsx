@echo off
setlocal
cd /d "%~dp0\.."

if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "release" rmdir /s /q "release"

echo Build output removed.
