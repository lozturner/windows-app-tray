@echo off
REM Laurence Trayhost — build, kill old, relaunch versioned binary
setlocal EnableDelayedExpansion
cd /d "%~dp0"

REM ----- pull version from source -----
for /f "tokens=3 delims= " %%a in ('findstr /c:"#define TRAYSYS_VERSION_A" src\main.cpp') do set "VERQ=%%a"
set "VER=%VERQ:"=%"
if "%VER%"=="" (
    echo ERROR: could not parse version from src\main.cpp
    exit /b 1
)
echo Building Laurence Trayhost v%VER% ...

REM ----- locate MSVC -----
set "VCVARS=C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvars64.bat"
if not exist "%VCVARS%" (
    echo ERROR: vcvars64.bat not found
    exit /b 1
)
call "%VCVARS%" >nul

set "OUT=C:\LaurenceTrayhost"
set "EXE=%OUT%\LaurenceTrayhost_v%VER%.exe"
if not exist "%OUT%" mkdir "%OUT%"

REM ----- kill any running instance BEFORE compiling so link.exe can overwrite -----
powershell -NoProfile -Command "Get-Process -Name 'LaurenceTrayhost*','TraySys*' -ErrorAction SilentlyContinue | Stop-Process -Force; Start-Sleep -Milliseconds 1200"

REM ----- compile resources (icons) -----
rc.exe /nologo /fo "%OUT%\app.res" src\app.rc
if errorlevel 1 (
    echo RC FAILED.
    exit /b 1
)

REM ----- compile -----
cl.exe /nologo /std:c++17 /O2 /EHsc /W3 /MT ^
    /D_WIN32_WINNT=0x0A00 /DWINVER=0x0A00 /DUNICODE /D_UNICODE /D_CRT_SECURE_NO_WARNINGS ^
    src\main.cpp "%OUT%\app.res" ^
    /Fo"%OUT%\\" /Fe"%EXE%" ^
    /link /SUBSYSTEM:WINDOWS ^
    user32.lib gdi32.lib shell32.lib shlwapi.lib ^
    psapi.lib dwmapi.lib ole32.lib ws2_32.lib advapi32.lib uuid.lib

if errorlevel 1 (
    echo BUILD FAILED.
    exit /b 1
)
del "%OUT%\main.obj" 2>nul

REM ----- launch new version -----
start "" "%EXE%"

echo.
echo OK.  Built and relaunched: %EXE%
echo Settings UI: http://127.0.0.1:8731/
endlocal
