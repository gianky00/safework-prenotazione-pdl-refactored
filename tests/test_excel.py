"""Test unitari per il modulo excel/processor.py."""

from unittest.mock import MagicMock, patch

import pytest

from src.config import Config
from src.excel.processor import ExcelProcessor
from src.models import CriticalConfigError


@pytest.fixture
def mock_openpyxl() -> MagicMock:
    """Fixture per mockare openpyxl."""
    with patch("src.excel.processor.openpyxl") as mock_xl:
        yield mock_xl


@pytest.fixture
def processor() -> ExcelProcessor:
    """Fixture per istanziare ExcelProcessor con un percorso fittizio (mockando exists per prevenire errori)."""
    with patch("src.excel.processor.os.path.exists") as mock_exists:
        mock_exists.return_value = True
        return ExcelProcessor("finto_config.xlsx")


@patch("src.excel.processor.os.path.exists")
def test_init(mock_exists: MagicMock) -> None:
    """Verifica l'inizializzazione corretta."""
    mock_exists.return_value = True
    proc = ExcelProcessor("finto_config.xlsx")
    assert proc.config_file_path == "finto_config.xlsx"
    assert proc.data_file_path is None


@patch("src.excel.processor.os.path.exists")
def test_init_not_found(mock_exists: MagicMock) -> None:
    """Verifica che l'inizializzazione sollevi errore se il file non esiste."""
    mock_exists.return_value = False
    with pytest.raises(CriticalConfigError):
        ExcelProcessor("inesistente.xlsx")


@patch("src.excel.processor.os.path.exists")
def test_leggi_parametri_configurazione(mock_exists: MagicMock, processor: ExcelProcessor, mock_openpyxl: MagicMock) -> None:
    """Testa il recupero dei parametri di configurazione da Excel."""
    mock_exists.return_value = True

    # Configura il comportamento dei mock per leggere le celle
    mock_wb = MagicMock()
    mock_openpyxl.load_workbook.return_value = mock_wb
    mock_sheet = MagicMock()
    mock_wb.__getitem__.return_value = mock_sheet

    values = {
        "B2": "C:\\dati_pdl.xlsx",
        "B3": "https://safework.isab.com/",
        "A3": "giancarlo",
    }
    mock_sheet.__getitem__.side_effect = lambda cell: MagicMock(value="pwd_segreta" if cell == "B3" else values.get(cell))

    url, user, pwd = processor.leggi_parametri_configurazione(interactive_pwd=False)
    assert url == "https://safework.isab.com/"
    assert user == "giancarlo"
    assert pwd == "pwd_segreta"
    assert processor.data_file_path == "C:\\dati_pdl.xlsx"


@patch("src.excel.processor.os.path.exists")
def test_leggi_parametri_configurazione_interactive(mock_exists: MagicMock, processor: ExcelProcessor, mock_openpyxl: MagicMock) -> None:
    """Testa il recupero parametri con password interattiva getpass."""
    mock_exists.return_value = True

    mock_wb = MagicMock()
    mock_openpyxl.load_workbook.return_value = mock_wb
    mock_sheet = MagicMock()
    mock_wb.__getitem__.return_value = mock_sheet

    mock_sheet.__getitem__.side_effect = lambda cell: MagicMock(
        value="C:\\dati_pdl.xlsx" if cell == "B2" else ("https://safework.isab.com/" if cell == "B3" else "giancarlo")
    )

    with patch("src.excel.processor.getpass.getpass") as mock_getpass:
        mock_getpass.return_value = "pwd_interattiva"
        url, user, pwd = processor.leggi_parametri_configurazione(interactive_pwd=True)
        assert pwd == "pwd_interattiva"


@patch("src.excel.processor.os.path.exists")
def test_leggi_parametri_configurazione_errors(mock_exists: MagicMock, processor: ExcelProcessor, mock_openpyxl: MagicMock) -> None:
    """Verifica la gestione degli errori durante la lettura dei parametri."""
    mock_exists.return_value = True

    mock_wb = MagicMock()
    mock_openpyxl.load_workbook.return_value = mock_wb
    mock_sheet = MagicMock()
    mock_wb.__getitem__.return_value = mock_sheet

    # Caso 1: Percorso dati vuoto
    mock_sheet.__getitem__.return_value = MagicMock(value=None)
    with pytest.raises(CriticalConfigError, match="Percorso file dati PdL non trovato"):
        processor.leggi_parametri_configurazione()

    # Caso 2: Username vuoto
    mock_sheet.__getitem__.side_effect = lambda cell: MagicMock(value="C:\\dati.xlsx" if cell == "B2" else None)
    with pytest.raises(CriticalConfigError, match="Username non trovato"):
        processor.leggi_parametri_configurazione()


@patch("src.excel.processor.os.path.exists")
def test_get_pdl_list_from_excel_empty(mock_exists: MagicMock, processor: ExcelProcessor) -> None:
    """Verifica che get_pdl_list_from_excel ritorni lista vuota se il file non esiste."""
    mock_exists.return_value = False
    assert processor.get_pdl_list_from_excel() == []


@patch("src.excel.processor.os.path.exists")
@patch("src.excel.processor.datetime")
def test_get_pdl_list_from_excel_bulk(mock_datetime: MagicMock, mock_exists: MagicMock, processor: ExcelProcessor, mock_openpyxl: MagicMock) -> None:
    """Verifica la lettura bulk dei PDL con righe valide e marker 'X'."""
    mock_exists.return_value = True
    processor.data_file_path = "finto_dati.xlsx"

    # Simula che oggi sia Lunedì (weekday = 0 -> colonna 'H')
    mock_date = MagicMock()
    mock_date.weekday.return_value = 0
    mock_datetime.now.return_value.date.return_value = mock_date

    mock_wb = MagicMock()
    mock_openpyxl.load_workbook.return_value = mock_wb
    mock_sheet = MagicMock()
    mock_wb.__getitem__.return_value = mock_sheet

    # Prepariamo una riga con marker 'X' a colonna giorno (indice 7)
    row_valida = [None] * 20
    row_valida[4] = "PDL-99"  # pdl
    row_valida[3] = "Nord"    # area
    row_valida[5] = "Impianto-X"  # impianto
    row_valida[6] = "Lavoro di test"  # descrizione
    row_valida[12] = "Da prenotare"  # stato_pdl_excel
    row_valida[16] = "Attivo"  # stato_attivita_excel
    row_valida[17] = "2026-05-18"  # data_controllo_excel
    row_valida[18] = "G. Allegretti"  # personale_excel
    row_valida[7] = "X"       # Lunedì marker

    # Riga non valida (senza marker)
    row_invalida = [None] * 20
    row_invalida[4] = "PDL-88"
    row_invalida[7] = ""

    mock_sheet.iter_rows.return_value = [row_valida, row_invalida]

    pdl_list = processor.get_pdl_list_from_excel()
    assert len(pdl_list) == 1
    assert pdl_list[0].pdl == "PDL-99"
    assert pdl_list[0].area == "Nord"


@patch("src.excel.processor.os.path.exists")
@patch("src.excel.processor.datetime")
def test_get_pdl_list_giorno_non_valido(mock_datetime: MagicMock, mock_exists: MagicMock, processor: ExcelProcessor) -> None:
    """Verifica che se non è un giorno programmato ritorni lista vuota."""
    mock_exists.return_value = True
    mock_date = MagicMock()
    mock_date.weekday.return_value = 6
    mock_datetime.now.return_value.date.return_value = mock_date

    assert processor.get_pdl_list_from_excel() == []


@patch("src.excel.processor.os.path.exists")
def test_aggiorna_stato_pdl_excel(mock_exists: MagicMock, processor: ExcelProcessor, mock_openpyxl: MagicMock) -> None:
    """Verifica l'aggiornamento dello stato PDL in Excel (scrittura su cella)."""
    mock_exists.return_value = True
    processor.data_file_path = "finto_dati.xlsx"

    mock_wb = MagicMock()
    mock_openpyxl.load_workbook.return_value = mock_wb
    mock_sheet = MagicMock()
    mock_wb.__getitem__.return_value = mock_sheet

    row_data = [None] * 20
    row_data[4] = "PDL-123"  # E (pdl)
    mock_sheet.iter_rows.return_value = [row_data]

    mock_cell = MagicMock()
    mock_sheet.cell.return_value = mock_cell

    success = processor.aggiorna_stato_pdl_excel("PDL-123", "Prenotato con successo")
    assert success is True
    assert mock_cell.value == "Prenotato con successo"
    mock_wb.save.assert_called_once()


@patch("src.excel.processor.os.path.exists")
def test_aggiorna_stato_pdl_non_trovato(mock_exists: MagicMock, processor: ExcelProcessor, mock_openpyxl: MagicMock) -> None:
    """Verifica che se il PDL non viene trovato non effettua scritture e ritorna False."""
    mock_exists.return_value = True
    processor.data_file_path = "finto_dati.xlsx"

    mock_wb = MagicMock()
    mock_openpyxl.load_workbook.return_value = mock_wb
    mock_sheet = MagicMock()
    mock_wb.__getitem__.return_value = mock_sheet

    row_data = [None] * 20
    row_data[4] = "PDL-ALTRO"
    mock_sheet.iter_rows.return_value = [row_data]

    success = processor.aggiorna_stato_pdl_excel("PDL-MIA-CHIAVE", "Prenotato")
    assert success is False


@patch("src.excel.processor.os.path.exists")
@patch("src.excel.processor.win32com.client")
def test_esegui_macro_sequenza_completa(mock_win32: MagicMock, mock_exists: MagicMock, processor: ExcelProcessor) -> None:
    """Verifica l'esecuzione corretta delle tre macro Excel tramite win32com."""
    mock_exists.return_value = True
    processor.data_file_path = "finto_dati.xlsx"

    mock_excel_app = MagicMock()
    mock_win32.Dispatch.return_value = mock_excel_app
    mock_wb = MagicMock()
    mock_excel_app.Workbooks.Open.return_value = mock_wb

    processor.esegui_macro_sequenza_completa()

    mock_excel_app.Run.assert_any_call(Config.MACRO_SEQ_1)
    mock_excel_app.Run.assert_any_call(Config.MACRO_SEQ_2)
    mock_excel_app.Run.assert_any_call(Config.MACRO_SEQ_3)

    mock_wb.Close.assert_called_once_with(SaveChanges=True)
    mock_excel_app.Quit.assert_called_once()
