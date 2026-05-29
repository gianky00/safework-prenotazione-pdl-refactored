"""Test unitari per il modulo main.py."""

import os
from unittest.mock import MagicMock, patch

import pytest

from src.main import PDLOrchestrator, StateManager, main
from src.models import PDLData


@pytest.fixture
def finto_stato_path() -> str:
    """Fixture per definire il percorso di un file di stato fittizio."""
    path = "finto_stato.json"
    yield path
    if os.path.exists(path):
        os.remove(path)


def test_state_manager_salva_carica(finto_stato_path: str) -> None:
    """Verifica il corretto salvataggio e ripristino dello stato tramite StateManager."""
    manager = StateManager(finto_stato_path)

    # Verifica stato iniziale vuoto
    idx, dati = manager.carica_stato()
    assert idx == -1
    assert dati == []

    # Salva uno stato
    pdl = PDLData(4, "12345", "Sud", "U100", "Lavoro", "Già Prenotato", "", "")
    manager.salva_stato(2, [pdl])

    # Ricarica e verifica
    idx, dati = manager.carica_stato()
    assert idx == 2
    assert len(dati) == 1
    assert dati[0].pdl == "12345"
    assert dati[0].stato_pdl_excel == "Già Prenotato"

    # Pulisce lo stato
    manager.pulisci_stato()
    idx, dati = manager.carica_stato()
    assert idx == -1
    assert dati == []


@patch("src.main.ExcelProcessor")
@patch("src.main.WebDriverManager")
@patch("src.main.StateManager")
def test_orchestrator_init(mock_state: MagicMock, mock_driver: MagicMock, mock_excel: MagicMock) -> None:
    """Verifica l'inizializzazione corretta di PDLOrchestrator."""
    orchestrator = PDLOrchestrator(dry_run=True, secure_pwd=False, headless=True)
    assert orchestrator.dry_run is True
    assert orchestrator.secure_pwd is False
    assert orchestrator.headless is True


@patch("src.main.ExcelProcessor")
@patch("src.main.WebDriverManager")
@patch("src.main.StateManager")
@patch("src.main.SafeWorkAutomator")
@patch("src.main.Progress")
def test_orchestrator_run_empty(
    mock_progress: MagicMock,
    mock_automator: MagicMock,
    mock_state_mgr: MagicMock,
    mock_drv_mgr: MagicMock,
    mock_excel_proc: MagicMock
) -> None:
    """Testa run() quando la lista dei PDL è vuota."""
    # Configura mock per il recupero dati
    mock_excel_instance = MagicMock()
    mock_excel_proc.return_value = mock_excel_instance
    mock_excel_instance.leggi_parametri_configurazione.return_value = (
        "https://safework.com", "utente", "password"
    )
    mock_excel_instance.get_credentials.return_value = ("utente", "password")
    mock_excel_instance.get_pdl_list_from_excel.return_value = []

    # Mock dello StateManager
    mock_state_instance = MagicMock()
    mock_state_mgr.return_value = mock_state_instance
    mock_state_instance.carica_stato.return_value = (-1, [])

    orchestrator = PDLOrchestrator(dry_run=True)
    orchestrator.run()

    # Dovrebbe aver eseguito la macro sequenza e letto i dati
    mock_excel_instance.esegui_macro_win32.assert_called_once()
    mock_excel_instance.get_pdl_list_from_excel.assert_called_once()

    # Non dovrebbe aver inizializzato il browser visto che la lista è vuota
    mock_drv_mgr.return_value.get_driver.assert_not_called()


@patch("src.main.ExcelProcessor")
@patch("src.main.WebDriverManager")
@patch("src.main.StateManager")
@patch("src.main.SafeWorkAutomator")
@patch("src.main.Progress")
def test_orchestrator_run_with_data(
    mock_progress: MagicMock,
    mock_automator: MagicMock,
    mock_state_mgr: MagicMock,
    mock_drv_mgr: MagicMock,
    mock_excel_proc: MagicMock
) -> None:
    """Testa il flusso di run() con dei PDL da elaborare."""
    mock_excel_instance = MagicMock()
    mock_excel_proc.return_value = mock_excel_instance
    mock_excel_instance.leggi_parametri_configurazione.return_value = (
        "https://safework.com", "utente", "password"
    )
    mock_excel_instance.get_credentials.return_value = ("utente", "password")

    pdl = PDLData(4, "12345", "Sud", "U100", "Lavoro", "", "", "")
    mock_excel_instance.get_pdl_list_from_excel.return_value = [pdl]

    mock_state_instance = MagicMock()
    mock_state_mgr.return_value = mock_state_instance
    mock_state_instance.carica_stato.return_value = (-1, [])

    # Mock per l'automatore
    mock_aut_instance = MagicMock()
    mock_automator.return_value = mock_aut_instance
    mock_aut_instance.prenota_singolo_pdl.return_value = "Prenotazione Eseguita"

    # Mock per il driver
    mock_drv_instance = MagicMock()
    mock_drv_mgr.return_value.get_driver.return_value = mock_drv_instance

    orchestrator = PDLOrchestrator(dry_run=False)
    orchestrator.run()

    # Dovrebbe aver eseguito il login
    mock_aut_instance.login_sito_safework.assert_called_once_with("https://safework.com", "utente", "password")

    # Dovrebbe aver aperto la pagina prenotazione
    mock_aut_instance.attiva_pagina_prenotazione.assert_called_once()

    # Dovrebbe aver elaborato il PDL
    mock_aut_instance.prenota_singolo_pdl.assert_called_once_with(pdl)

    # Dovrebbe aver salvato lo stato in Excel
    mock_excel_instance.aggiorna_stato_pdl_excel.assert_called_once_with("12345", "Prenotazione Eseguita")

    # Dovrebbe aver pulito lo stato JSON finale
    mock_state_instance.pulisci_stato.assert_called_once()


@patch("src.main.argparse.ArgumentParser")
@patch("src.main.PDLOrchestrator")
def test_main_cli(mock_orchestrator: MagicMock, mock_parser: MagicMock) -> None:
    """Verifica il parsing CLI e l'invocazione dell'orchestratore."""
    mock_args = MagicMock()
    mock_args.dry_run = True
    mock_args.secure = False
    mock_args.headless = True
    mock_parser.return_value.parse_args.return_value = mock_args

    main()

    # Dovrebbe aver istanziato l'orchestratore con i parametri della riga di comando
    mock_orchestrator.assert_called_once_with(dry_run=True, secure_pwd=False, headless=True)
    mock_orchestrator.return_value.run.assert_called_once()
