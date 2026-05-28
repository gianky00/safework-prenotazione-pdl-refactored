@echo off
title SafeWork PDL - Avvio Manuale
cls

echo.
echo ============================================================
echo        SAFEWORK PDL - SELEZIONE MODALITA PRENOTAZIONE
echo ============================================================
echo.
echo [1] OGGI PER DOMANI (Standard - B6=SI)
echo [2] OGGI PER OGGI   (Manuale  - B6=NO)
echo.
echo ============================================================
echo.

set /p choice="Seleziona la modalita (1 o 2) [Default=1]: "

if "%choice%"=="2" (
    echo.
    echo Avvio in modalita OGGI PER OGGI...
    python -m src.main --today
) else (
    echo.
    echo Avvio in modalita OGGI PER DOMANI...
    python -m src.main
)

pause
