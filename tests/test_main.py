"""Test unitari per il modulo main.py."""

import os
import json
from unittest.mock import MagicMock, patch
import pytest

from src.main import PDLOrchestrator, StateManager
from src.models import PDLData


@pytest.fixture
def finto_stato_path() -> str:
    """Fixture per un file di stato fittizio."""
    path = "finto_stato.json"
    yield path
    if os.path.exists(path):
        os.remove(path)


def test_state_manager_salva_carica(finto_stato_path: str) -> None:
    """Testa salvataggio e caricamento stato."""
    manager = StateManager(finto_stato_path)
    
    pdl = PDLData(1, "123", "Area", "Imp", "Desc", "StatoExcel", "", "")
    manager.salva_stato(1, [pdl])
    
    idx, dati = manager.carica_stato()
    assert idx == 1
    assert len(dati) == 1
    assert dati[0].pdl == "123"


@patch("src.main.ExcelProcessor")
@patch("src.main.WebDriverManager")
@patch("src.main.StateManager")
@patch("src.main.SafeWorkAutomator")
@patch("src.main.PrinterManager")
@patch("src.main.EmailManager")
def test_orchestrator_run_flow(
    mock_email: MagicMock,
    mock_printer: MagicMock,
    mock_automator: MagicMock,
    mock_state: MagicMock,
    mock_driver: MagicMock,
    mock_excel: MagicMock
) -> None:
    """Testa il flusso principale dell'orchestratore."""
    # Configura Excel
    mock_excel_instance = mock_excel.return_value
    mock_excel_instance.get_website_url.return_value = "http://test.com"
    mock_excel_instance.get_credentials.return_value = ("user", "pwd")
    pdl = PDLData(1, "123", "Area", "Imp", "Desc", "", "", "")
    mock_excel_instance.get_pdl_list_from_excel.return_value = [pdl]
    
    # Configura StateManager
    mock_state.return_value.carica_stato.return_value = (-1, [])
    
    # Configura Automator
    mock_automator.return_value.process_pdl.return_value = "SUCCESSO"
    
    orchestrator = PDLOrchestrator(dry_run=False)
    orchestrator.run()
    
    # Verifica chiamate principali
    mock_excel_instance.run_pdl_macros.assert_called()
    mock_automator.return_value.login.assert_called()
    mock_automator.return_value.process_pdl.assert_called_with(pdl)
