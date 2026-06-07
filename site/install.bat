@echo off
setlocal
REM ===========================================================================
REM  Rostrum - one-click Windows installer (Stage A.5)
REM ---------------------------------------------------------------------------
REM  Registers the hosted Rostrum add-in with desktop Word for the CURRENT USER
REM  by writing the single per-user "developer sideload" registry value Word
REM  reads at launch. No admin rights, no Node, no dev server.
REM
REM  Mechanism: HKCU\...\WEF\Developer\<add-in Id> = <path to a local manifest>.
REM  This is exactly what office-addin-debugging writes under the hood.
REM
REM  Re-running this file is also the UPDATE path: it re-downloads the current
REM  production manifest, so a future ribbon change reaches you on the next run.
REM
REM  ADDIN_ID MUST equal the <Id> of the deployed prod manifest. A drift test
REM  (__tests__/installerBat.test.ts) fails the build if they fall out of sync.
REM ===========================================================================

set "MANIFEST_DIR=%LOCALAPPDATA%\Rostrum"
set "MANIFEST_PATH=%MANIFEST_DIR%\manifest.xml"
set "MANIFEST_URL=https://andrewtjin.github.io/rostrum/manifest.xml"
set "ADDIN_ID=ea3fb238-6832-4f91-9654-b9e7ef24d926"
set "WEF_KEY=HKCU\Software\Microsoft\Office\16.0\WEF\Developer"

echo.
echo  Installing Rostrum for Word...
echo.

REM curl.exe ships with Windows 10 1803+ (universal in 2026). If it is somehow
REM missing, send the student to the manual steps rather than failing cryptically.
where curl >nul 2>&1
if errorlevel 1 (
  echo  This installer needs Windows 10 version 1803 or newer.
  echo  Please use the manual install steps on the Rostrum website instead.
  echo.
  pause
  exit /b 1
)

REM Create the per-user install folder. Guard the result so a filesystem failure
REM reports the RIGHT cause instead of the misleading "check your internet" below.
if not exist "%MANIFEST_DIR%" mkdir "%MANIFEST_DIR%" 2>nul
if not exist "%MANIFEST_DIR%" (
  echo  Could not create Rostrum's install folder in your user profile.
  echo  Make sure your account can write to its AppData folder, then retry.
  echo.
  pause
  exit /b 1
)

REM Fetch the LIVE production manifest. -f makes HTTP errors (404 etc.) fail the
REM command; -sS stays quiet but still prints a real error if one occurs.
curl -fsS -o "%MANIFEST_PATH%" "%MANIFEST_URL%"
if errorlevel 1 (
  echo  Could not download Rostrum. Check your internet connection and try again.
  echo  If your school network blocks this, use the manual steps on the website.
  echo.
  pause
  exit /b 1
)

REM Sanity-check the download. curl -f catches HTTP errors but NOT a captive-portal
REM sign-in page (HTTP 200 + HTML) or a connection that drops mid-download. Requiring
REM the manifest's CLOSING tag proves we got a COMPLETE Office manifest - not a partial
REM file or a login page, either of which would register and then silently fail in Word.
findstr /i /c:"</OfficeApp>" "%MANIFEST_PATH%" >nul 2>&1
if errorlevel 1 (
  echo  Rostrum did not download correctly - your network may have returned a
  echo  sign-in page, or the connection dropped partway. Connect to a normal
  echo  network and try again, or use the manual steps on the website.
  echo.
  del "%MANIFEST_PATH%" >nul 2>&1
  pause
  exit /b 1
)

REM Point Word's per-user developer-sideload key at the local manifest copy.
REM /f overwrites any existing entry, so re-running is safe and idempotent.
reg add "%WEF_KEY%" /v "%ADDIN_ID%" /t REG_SZ /d "%MANIFEST_PATH%" /f >nul
if errorlevel 1 (
  echo  Could not register Rostrum with Word. On a school-managed laptop the
  echo  registry may be locked by policy - use the manual steps, or ask IT.
  echo.
  pause
  exit /b 1
)

echo  Rostrum is installed for your account.
echo.

REM The WEF key is only read when Word LAUNCHES. If Word is already open the add-in
REM won't appear until a full restart - so detect a running Word and say so, otherwise
REM a student sees nothing and assumes it failed. POLARITY NOTE: `find` exits 0 when it
REM MATCHES (Word running) and 1 when not, so `if errorlevel 1` is the NOT-running branch.
tasklist /fi "imagename eq winword.exe" 2>nul | find /i "winword.exe" >nul
if errorlevel 1 (
  echo  Next: open Word and look for the "Rostrum" tab on the ribbon
  echo  ^(or Home ^> Add-ins ^> Developer Add-ins^).
) else (
  echo  Word is open right now. Close it COMPLETELY, reopen it, and look for the
  echo  "Rostrum" tab on the ribbon ^(or Home ^> Add-ins ^> Developer Add-ins^).
)
echo.
echo  If the "Rostrum" tab never appears, your Word may be a Microsoft Store
echo  install or managed by your school - use the manual steps on the website.
echo.
pause
