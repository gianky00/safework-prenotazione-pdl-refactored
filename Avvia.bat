@echo off
title SafeWork Prenotazione PDL - Enterprise Edition
setlocal

:: Forza l'encoding UTF-8 per Python
set PYTHONUTF8=1

:: Percorso della directory dello script
set "BASE_DIR=%~dp0"
cd /d "%BASE_DIR%"

:: Esecuzione del modulo tramite il package src
:: Nota: assume che le dipendenze siano installate nell'ambiente python corrente
python -m src.main %*

if %ERRORLEVEL% neq 0 (
    echo.
    echo ❌ Errore durante l'esecuzione dell'automazione.
    :: Pausa solo se avviato interattivamente in una console cmd visibile
    echo %cmdcmdline% | findstr /i "cmd.exe" >nul
    if %errorlevel% equ 0 pause
) else (
    echo.
    echo ✅ Processo completato con successo.
    :: Attesa non bloccante sicura per l'Utilità di Pianificazione
    ping 127.0.0.1 -n 6 >nul
)

endlocal
