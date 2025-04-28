@echo off
setlocal enabledelayedexpansion

REM --- Pfad zur R-Script-Datei (relativ zum aktuellen Ordner) ---
set "RScriptPath=NFM_DASHBOARD.R"

REM --- Optional: Zielversion festlegen (maximal zulässige Version) ---
set "TARGETVERSION=R-4.5.0"

if defined TARGETVERSION (
    for /f "tokens=2 delims=-" %%T in ("%TARGETVERSION%") do (
        set "targetVer=%%T"
    )
    for /f "tokens=1-3 delims=." %%x in ("!targetVer!") do (
        set "targetMajor=%%x"
        set "targetMinor=%%y"
        set "targetPatch=%%z"
    )
    set /a targetNum=targetMajor*10000 + targetMinor*100 + targetPatch
)

REM --- In das Verzeichnis der BAT-Datei wechseln ---
pushd %~dp0
echo Aktuelles Verzeichnis: %CD%

REM --- Standardpfad für R durchsuchen ---
set "RBasePath=C:\Program Files\R"
set "RScriptExe="
set "RVersion="
set "maxNum=0"
for /d %%V in ("%RBasePath%\R-*") do (
    if exist "%%V\bin\Rscript.exe" (
         for /f "tokens=2 delims=-" %%A in ("%%~nxV") do (
             for /f "tokens=1-3 delims=." %%x in ("%%A") do (
                 set "major=%%x"
                 set "minor=%%y"
                 set "patch=%%z"
             )
             set /a currentNum=major*10000 + minor*100 + patch
             if defined TARGETVERSION (
                 if !currentNum! leq !targetNum! if !currentNum! gtr !maxNum! (
                     set "maxNum=!currentNum!"
                     set "RScriptExe=%%V\bin\Rscript.exe"
                     set "RVersion=%%~nxV"
                 )
             ) else (
                 if !currentNum! gtr !maxNum! (
                     set "maxNum=!currentNum!"
                     set "RScriptExe=%%V\bin\Rscript.exe"
                     set "RVersion=%%~nxV"
                 )
             )
         )
    )
)

REM --- Falls im Standardpfad nichts gefunden wurde, alternative Suche im Benutzerpfad ---
if "!RScriptExe!"=="" (
    set "AltPath=C:\Users\%USERNAME%\AppData\Local\Programs\R"
    set "maxNum=0"
    for /d %%V in ("!AltPath!\R-*") do (
        if exist "%%V\bin\Rscript.exe" (
             for /f "tokens=2 delims=-" %%A in ("%%~nxV") do (
                 for /f "tokens=1-3 delims=." %%x in ("%%A") do (
                     set "major=%%x"
                     set "minor=%%y"
                     set "patch=%%z"
                 )
                 set /a currentNum=major*10000 + minor*100 + patch
                 if defined TARGETVERSION (
                     if !currentNum! leq !targetNum! if !currentNum! gtr !maxNum! (
                         set "maxNum=!currentNum!"
                         set "RScriptExe=%%V\bin\Rscript.exe"
                         set "RVersion=%%~nxV"
                     )
                 ) else (
                     if !currentNum! gtr !maxNum! (
                         set "maxNum=!currentNum!"
                         set "RScriptExe=%%V\bin\Rscript.exe"
                         set "RVersion=%%~nxV"
                     )
                 )
             )
        )
    )
)

:found
if "!RScriptExe!"=="" (
    echo Fehler: Keine passende Rscript.exe gefunden! Stelle sicher, dass R installiert ist.
    pause
    popd
    exit /b 1
)

echo Gefundene R-Version: !RVersion!
echo Verwendeter Rscript-Pfad: !RScriptExe!

REM --- R-Script ausführen ---
"!RScriptExe!" "!RScriptPath!"

popd
endlocal
