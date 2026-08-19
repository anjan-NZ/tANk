@echo off
setlocal

rem Removes tANk from your copy of Office. Double click this file.

set "ADDINID=7f3c1a54-9e2b-4d61-8a77-2c5b9e0f41d3"
set "DIR=%LOCALAPPDATA%\tANk"
set "WEF=%LOCALAPPDATA%\Microsoft\Office\16.0\Wef"
set "SYS=%SystemRoot%\System32"

echo.
echo   Removing tANk...
echo.

for %%P in (EXCEL.EXE WINWORD.EXE POWERPNT.EXE) do (
  "%SYS%\tasklist.exe" /FI "IMAGENAME eq %%P" /NH 2>nul | "%SYS%\find.exe" /I "%%P" >nul && (
    echo   Close Excel, Word and PowerPoint first, then run this again.
    echo.
    pause
    exit /b 1
  )
)

"%SYS%\reg.exe" delete "HKCU\SOFTWARE\Microsoft\Office\16.0\WEF\Developer" /v "%ADDINID%" /f >nul 2>nul

rem Guarded, so an unusual environment cannot widen this into something else.
if defined LOCALAPPDATA if exist "%DIR%\manifest.xml" rd /s /q "%DIR%" 2>nul

for %%C in (AppCommands AddinInfo) do (
  if exist "%WEF%\%%C" rd /s /q "%WEF%\%%C" 2>nul
)

echo   Done. The button will be gone next time Office starts.
echo.
echo   One thing left: any API keys you pasted are stored by the pane itself, inside
echo   Office's web storage, which this script cannot pick apart without wiping the
echo   saved data of every other add-in on this PC.
echo.
echo   If you still have tANk installed, open it and press Settings, then Erase
echo   everything. That clears the keys and every setting straight away.
echo.
pause
