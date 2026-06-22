@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo.
echo ===================================================
echo  ExoSync ^| Sync rapide TrainSystem
echo ===================================================
echo.

:: --- 1. Sauvegarder et capturer les fichiers Excel ouverts ---
set TEMP_LISTE=%TEMP%\excel_ouverts.txt
if exist "%TEMP_LISTE%" del "%TEMP_LISTE%"

echo [1/4] Sauvegarde et detection des fichiers Excel ouverts...
powershell -NoProfile -Command ^
  "try { $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application'); $excel.DisplayAlerts = $false; $saved = 0; $skipped = 0; $liste = [System.Collections.Generic.List[string]]::new(); foreach ($wb in $excel.Workbooks) { if ($wb.Path -ne '') { $wb.Save(); $saved++; $liste.Add($wb.FullName); Write-Host ('   Sauvegarde : ' + $wb.Name) } else { $skipped++ } }; if ($liste.Count -gt 0) { [System.IO.File]::WriteAllLines('%TEMP_LISTE%', $liste) }; Write-Host ('   ' + $saved + ' fichier(s) sauvegardes, ' + $skipped + ' ignores (jamais enregistres).') } catch { Write-Host '   Aucune instance Excel active detectee.' }" ^
  2>nul

if exist "%TEMP_LISTE%" (
    echo    Fichiers captures :
    for /f "usebackq tokens=*" %%i in ("%TEMP_LISTE%") do echo      - %%i
) else (
    echo    Aucun fichier Excel ouvert detecte.
)

:: --- 2. Fermer Excel ---
echo.
echo [2/4] Fermeture d'Excel...
taskkill /IM excel.exe /F >nul 2>&1
if %errorlevel% == 0 (
    echo    Excel ferme.
) else (
    echo    Excel n'etait pas ouvert.
)
timeout /t 2 /nobreak >nul

:: --- 3. Lancer la synchronisation ---
echo.
echo [3/4] Synchronisation en cours...
echo ---------------------------------------------------
python -m src sync --dir projet_TrainSystem
set SYNC_CODE=%errorlevel%
echo ---------------------------------------------------

if %SYNC_CODE% neq 0 (
    echo.
    echo [ERREUR] La synchronisation a echoue ^(code %SYNC_CODE%^).
    echo Appuyez sur une touche pour continuer quand meme...
    pause >nul
)

:: --- 4. Reouvrir les fichiers ---
echo.
echo [4/4] Reouverture des fichiers Excel...
if exist "%TEMP_LISTE%" (
    for /f "usebackq tokens=*" %%i in ("%TEMP_LISTE%") do (
        if exist "%%i" (
            echo    Ouverture : %%i
            start "" "%%i"
            timeout /t 1 /nobreak >nul
        ) else (
            echo    [WARN] Introuvable : %%i
        )
    )
    del "%TEMP_LISTE%"
) else (
    echo    Aucun fichier a reouvrir.
)

echo.
echo ===================================================
echo  Sync terminee. Appuyez sur une touche pour fermer.
echo ===================================================
pause >nul
