"""Test unitari per il modulo config.py e models.py."""

from src.config import Config
from src.models import AutomationError, CriticalConfigError, PDLData


def test_config_values() -> None:
    """Verifica che tutte le costanti di configurazione di base siano impostate correttamente."""
    assert Config.DEFAULT_URL_SITO == "https://safework.isab.com/"
    assert Config.NOME_FOGLIO_DATI_PDL == "Riepilogo"
    assert Config.EXCEL_SHEET_CREDENTIALS == "credenziali"
    assert Config.EXCEL_SHEET_PERCORSI == "percorsi"


def test_models_pdl_data() -> None:
    """Verifica la creazione e validità del modello dati PDLData."""
    pdl = PDLData(
        riga_excel_debug=4,
        pdl="12345",
        area="Sud",
        impianto="U100",
        descrizione="Test impianto",
        stato_pdl_excel="Da prenotare",
        stato_attivita_excel="Attivo",
        data_controllo_excel="2026-05-18",
        personale_excel="G. Allegretti"
    )
    assert pdl.riga_excel_debug == 4
    assert pdl.pdl == "12345"
    assert pdl.area == "Sud"
    assert pdl.impianto == "U100"
    assert pdl.descrizione == "Test impianto"
    assert pdl.stato_pdl_excel == "Da prenotare"
    assert pdl.stato_attivita_excel == "Attivo"
    assert pdl.data_controllo_excel == "2026-05-18"
    assert pdl.personale_excel == "G. Allegretti"


def test_exceptions() -> None:
    """Verifica il comportamento delle eccezioni custom."""
    err = AutomationError("Errore automazione")
    assert str(err) == "Errore automazione"

    cfg_err = CriticalConfigError("Errore configurazione")
    assert str(cfg_err) == "Errore configurazione"
