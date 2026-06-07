@echo off
setlocal
REM ===========================================================================
REM  Rostrum - uninstaller (Stage A.5)
REM ---------------------------------------------------------------------------
REM  Reverses install.bat: removes the per-user developer-sideload registry
REM  entry and deletes the local manifest copy. No admin rights needed.
REM
REM  It does NOT clear Office's own web add-in cache - if the Rostrum tab still
REM  lingers after restarting Word, see the Uninstall section on the website.
REM ===========================================================================

set "ADDIN_ID=ea3fb238-6832-4f91-9654-b9e7ef24d926"
set "WEF_KEY=HKCU\Software\Microsoft\Office\16.0\WEF\Developer"

REM Both are best-effort: succeed silently whether or not the entry/folder exist.
reg delete "%WEF_KEY%" /v "%ADDIN_ID%" /f >nul 2>&1
rmdir /s /q "%LOCALAPPDATA%\Rostrum" >nul 2>&1

echo.
echo  Rostrum has been removed. Restart Word to finish.
echo  If the "Rostrum" tab still lingers, clear Office's web add-in cache -
echo  see the Uninstall section on the Rostrum website.
echo.
pause
