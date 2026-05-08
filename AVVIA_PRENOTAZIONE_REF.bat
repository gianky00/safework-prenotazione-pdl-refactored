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
python -m src.prenotazione_pdl.main %*

if %ERRORLEVEL% neq 0 (
    echo.
    echo ❌ Errore durante l'esecuzione dell'automazione.
    pause
) else (
    echo.
    echo ✅ Processo completato con successo.
    timeout /t 5
)

endlocal
