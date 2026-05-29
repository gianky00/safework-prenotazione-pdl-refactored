"""Test unitari per il modulo excel/processor.py."""

from unittest.mock import MagicMock, patch
from typing import Any
import pytest

from src.excel.processor import ExcelProcessor
from src.models import CriticalConfigError, PDLData


@pytest.fixture
def processor() -> ExcelProcessor:
    """Fixture per istanziare ExcelProcessor con un percorso fittizio."""
    with patch("src.excel.processor.os.path.exists") as mock_exists:
        mock_exists.return_value = True
        return ExcelProcessor("finto_config.xlsx")


def test_init(processor: ExcelProcessor) -> None:
    """Verifica l'inizializzazione corretta."""
    assert processor.config_file_path == "finto_config.xlsx"
    assert processor.data_file_path is None


@patch("src.excel.processor.os.path.exists")
def test_init_not_found(mock_exists: MagicMock) -> None:
    """Verifica che l'inizializzazione sollevi errore se il file non esiste."""
    mock_exists.return_value = False
    with pytest.raises(CriticalConfigError):
        ExcelProcessor("inesistente.xlsx")


def test_get_website_url(processor: ExcelProcessor) -> None:
    """Testa il recupero dell'URL del sito."""
    with patch.object(processor, "_leggi_cella") as mock_leggi:
        mock_leggi.return_value = "https://test.com"
        assert processor.get_website_url() == "https://test.com"


def test_get_website_url_fallback(processor: ExcelProcessor) -> None:
    """Testa il fallback dell'URL in caso di valore non valido."""
    with patch.object(processor, "_leggi_cella") as mock_leggi:
        mock_leggi.return_value = None
        # Verifica che ritorni un URL di default (non hardcoded qui per flessibilità)
        url = processor.get_website_url()
        assert url.startswith("http")


def test_get_pdl_data_file_path(processor: ExcelProcessor) -> None:
    """Testa il recupero del percorso file dati."""
    with patch.object(processor, "_leggi_cella") as mock_leggi:
        mock_leggi.return_value = "C:\\dati.xlsx"
        assert processor.get_pdl_data_file_path() == "C:\\dati.xlsx"


def test_get_credentials(processor: ExcelProcessor) -> None:
    """Testa il recupero delle credenziali."""
    with patch.object(processor, "_leggi_cella") as mock_leggi:
        mock_leggi.side_effect = ["utente", "password"]
        user, pwd = processor.get_credentials()
        assert user == "utente"
        assert pwd == "password"


@patch("src.excel.processor.getpass.getpass")
def test_get_credentials_interactive(mock_getpass: MagicMock, processor: ExcelProcessor) -> None:
    """Testa il recupero credenziali in modalità interattiva."""
    with patch.object(processor, "_leggi_cella") as mock_leggi:
        mock_leggi.return_value = "utente"
        mock_getpass.return_value = "pwd_interattiva"
        user, pwd = processor.get_credentials(interactive_pwd=True)
        assert user == "utente"
        assert pwd == "pwd_interattiva"


@patch("src.excel.processor.os.path.exists")
def test_get_pdl_list_from_excel_empty_file(mock_exists: MagicMock, processor: ExcelProcessor) -> None:
    """Verifica lista vuota se il file dati non esiste."""
    mock_exists.return_value = False
    processor.data_file_path = "inesistente.xlsx"
    assert processor.get_pdl_list_from_excel() == []


@patch("src.excel.processor.WIN32COM_AVAILABLE", True)
def test_run_pdl_macros_success(processor: ExcelProcessor) -> None:
    """Verifica il successo dell'esecuzione macro."""
    with patch.object(processor, "get_pdl_data_file_path") as mock_path:
        mock_path.return_value = "dati.xlsx"
        # Impostiamo il percorso nel processore per superare il check os.path.exists
        processor.data_file_path = "dati.xlsx"
        with patch("src.excel.processor.os.path.exists") as mock_exists:
            mock_exists.return_value = True
            with patch.object(processor, "_esegui_sessione_macro") as mock_session:
                mock_session.return_value = [PDLData(1, "123", "Area", "Imp", "Desc", "", "", "")]
                assert processor.run_pdl_macros() is True
                assert len(processor._cached_pdl_list) == 1


def test_formatta_valore_cella(processor: ExcelProcessor) -> None:
    """Testa la formattazione dei valori delle celle."""
    assert processor._formatta_valore_cella(123.0, "pdl") == "123"
    assert processor._formatta_valore_cella("  test  ", "area") == "test"
    assert processor._formatta_valore_cella(None, "impianto") == ""
