"""Test di fumo (smoke tests) per verificare la coerenza del modulo prenotazione_pdl."""

import os

from src.config import Config
from src.excel.processor import ExcelProcessor
from src.main import PDLOrchestrator


def test_config_initialization() -> None:
    """Verifica che la configurazione sia importabile e coerente."""
    assert Config.SCRIPT_DIR is not None
    assert isinstance(Config.SCRIPT_DIR, str)
    assert os.path.exists(Config.SCRIPT_DIR)
    assert Config.EXCEL_FILE_CONFIG_NAME == "parametri prenotazione pdl.xlsx"


def test_excel_processor_import() -> None:
    """Verifica che l'ExcelProcessor sia importabile ed esponibile."""
    config_path = os.path.join(Config.SCRIPT_DIR, Config.EXCEL_FILE_CONFIG_NAME)
    processor = ExcelProcessor(config_path)
    assert processor.config_file_path == config_path


def test_orchestrator_initialization() -> None:
    """Verifica che l'orchestratore possa essere istanziato in modalità dry-run."""
    # Verifichiamo l'inizializzazione dell'orchestratore in modalità headless
    orchestrator = PDLOrchestrator(dry_run=True, secure_pwd=False, headless=True)
    assert orchestrator.dry_run is True
    assert orchestrator.headless is True
    assert orchestrator.driver_manager is not None
    # Chiudiamo subito il driver manager associato
    orchestrator.driver_manager.quit_driver()
