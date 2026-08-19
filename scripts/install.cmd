@echo off
setlocal

rem tANk installer. Double click this file.
rem Nothing is compiled or executed from a temporary folder, so Smart App Control and
rem application control policies have nothing to block. It copies one file and writes
rem one registry value, both inside your own user profile.

set "BASE=__BASE_URL__"
set "ADDINID=7f3c1a54-9e2b-4d61-8a77-2c5b9e0f41d3"
set "DIR=%LOCALAPPDATA%\tANk"
set "WEF=%LOCALAPPDATA%\Microsoft\Office\16.0\Wef"
set "SYS=%SystemRoot%\System32"

echo.
echo   tANk - AI chat inside Excel, Word and PowerPoint
echo   ------------------------------------------------
echo.

rem Office has to be closed or it will not notice the new add-in. Full paths, because
rem a machine with unix tools ahead of System32 on PATH would run the wrong find.
for %%P in (EXCEL.EXE WINWORD.EXE POWERPNT.EXE) do (
  "%SYS%\tasklist.exe" /FI "IMAGENAME eq %%P" /NH 2>nul | "%SYS%\find.exe" /I "%%P" >nul && (
    echo   Close Excel, Word and PowerPoint first, then run this again.
    echo.
    pause
    exit /b 1
  )
)

echo   Downloading the add-in description...
if not exist "%DIR%" mkdir "%DIR%"
"%SYS%\curl.exe" -fsSL -o "%DIR%\manifest.xml" "%BASE%/manifest.xml"
if errorlevel 1 (
  echo   Could not reach %BASE% - check your internet connection.
  echo.
  pause
  exit /b 1
)

echo   Registering with Office...
"%SYS%\reg.exe" add "HKCU\SOFTWARE\Microsoft\Office\16.0\WEF\Developer" /v "%ADDINID%" /t REG_SZ /d "%DIR%\manifest.xml" /f >nul

rem Office caches the ribbon buttons. Only these two folders are cleared and Office
rem rebuilds both on its next start. The neighbouring webview2 folder, where every
rem add-in on this PC keeps its saved data, is never touched.
for %%C in (AppCommands AddinInfo) do (
  if exist "%WEF%\%%C" rd /s /q "%WEF%\%%C" 2>nul
)

echo.
echo   Done.
echo.
echo   Open Excel and look for tANk at the right of the Home tab. If it is not there
echo   yet, use Add-ins on the same tab and pick tANk under Developer Add-ins.
echo   First time: open Settings in the pane and paste a free API key.
echo.
echo   To remove it later, run uninstall.cmd from the same place.
echo.
pause
