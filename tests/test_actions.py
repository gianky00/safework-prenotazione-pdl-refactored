"""Test unitari per il modulo automation/actions.py."""

from unittest.mock import MagicMock, patch
import pytest
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from src.automation.actions import SafeWorkAutomator
from src.config import Config
from src.models import AutomationError, PDLData, TimeoutAlertError


@pytest.fixture
def mock_driver() -> MagicMock:
    """Fixture per creare un WebDriver simulato."""
    return MagicMock()


@pytest.fixture
def automator(mock_driver: MagicMock) -> SafeWorkAutomator:
    """Fixture per istanziare SafeWorkAutomator."""
    return SafeWorkAutomator(mock_driver)


def test_automator_init(automator: SafeWorkAutomator, mock_driver: MagicMock) -> None:
    """Verifica l'inizializzazione corretta."""
    assert automator.driver == mock_driver
    assert automator.current_pdl_context is None


@patch("src.automation.actions.WebDriverWait")
def test_wait_and_find_success(mock_wait: MagicMock, automator: SafeWorkAutomator) -> None:
    """Testa la ricerca di elementi con successo tramite WebDriverWait."""
    mock_wait_instance = MagicMock()
    mock_wait.return_value = mock_wait_instance

    mock_element = MagicMock()
    mock_wait_instance.until.return_value = mock_element

    selectors = [(By.ID, "finto_id")]
    el = automator._wait_and_find(selectors)

    assert el == mock_element
    mock_wait.assert_called_once_with(automator.driver, Config.SELENIUM_TIMEOUT_SHORT, poll_frequency=0.2)


@patch("src.automation.actions.WebDriverWait")
def test_wait_and_find_failure(mock_wait: MagicMock, automator: SafeWorkAutomator) -> None:
    """Verifica che se la ricerca elementi fallisce su tutti i selettori venga sollevata eccezione."""
    mock_wait_instance = MagicMock()
    mock_wait.return_value = mock_wait_instance
    mock_wait_instance.until.side_effect = TimeoutException("Elemento non trovato")

    selectors = [(By.ID, "inesistente_1"), (By.ID, "inesistente_2")]
    with pytest.raises(AutomationError, match="Impossibile trovare l'elemento"):
        automator._wait_and_find(selectors)


@patch("src.automation.actions.WebDriverWait")
def test_click_element_success(mock_wait: MagicMock, automator: SafeWorkAutomator) -> None:
    """Verifica il click corretto su un elemento trovato."""
    mock_wait_instance = MagicMock()
    mock_wait.return_value = mock_wait_instance

    mock_element = MagicMock()
    mock_wait_instance.until.return_value = mock_element

    automator._click_element([(By.ID, "btn")], "Bottone Test")
    mock_element.click.assert_called_once()


@patch("src.automation.actions.WebDriverWait")
def test_inserisci_testo_success(mock_wait: MagicMock, automator: SafeWorkAutomator) -> None:
    """Verifica l'inserimento corretto di testo nei campi input."""
    mock_wait_instance = MagicMock()
    mock_wait.return_value = mock_wait_instance

    mock_element = MagicMock()
    mock_wait_instance.until.return_value = mock_element

    automator._inserisci_testo([(By.ID, "input")], "Ciao Mondo", "Input Test")
    mock_element.clear.assert_called_once()
    mock_element.send_keys.assert_called_once_with("Ciao Mondo")


@patch("src.automation.actions.WebDriverWait")
@patch("src.automation.actions.time.sleep")
def test_login_sito_safework(mock_sleep: MagicMock, mock_wait: MagicMock, automator: SafeWorkAutomator) -> None:
    """Testa l'intero processo di login sul portale."""
    mock_wait_instance = MagicMock()
    mock_wait.return_value = mock_wait_instance

    mock_element = MagicMock()
    mock_wait_instance.until.return_value = mock_element

    # Configura il driver per simulare il caricamento
    automator.driver.current_url = "https://safework.isab.com/home"

    # Esegue il login
    automator.login_sito_safework("https://safework.isab.com/", "mio_utente", "mia_pwd")

    # Verifica navigazione iniziale
    automator.driver.get.assert_called_once_with("https://safework.isab.com/")


@patch("src.automation.actions.WebDriverWait")
def test_attiva_pagina_prenotazione(mock_wait: MagicMock, automator: SafeWorkAutomator) -> None:
    """Testa l'apertura della pagina di prenotazione PDL."""
    mock_wait_instance = MagicMock()
    mock_wait.return_value = mock_wait_instance

    mock_element = MagicMock()
    mock_wait_instance.until.return_value = mock_element

    automator.attiva_pagina_prenotazione()

    # Dovrebbe cliccare su HOME_BUTTON e poi su LINK_PRENOTAZIONE_PDL
    assert mock_element.click.call_count >= 2


@patch("src.automation.actions.WebDriverWait")
@patch("src.automation.actions.time.sleep")
def test_prenota_singolo_pdl_already_booked(mock_sleep: MagicMock, mock_wait: MagicMock, automator: SafeWorkAutomator) -> None:
    """Testa la logica se il PDL è già prenotato."""
    mock_wait_instance = MagicMock()
    mock_wait.return_value = mock_wait_instance

    mock_element = MagicMock()
    mock_wait_instance.until.return_value = mock_element

    pdl = PDLData(
        riga_excel_debug=4,
        pdl="112233",
        area="Sud",
        impianto="U100",
        descrizione="Test",
        stato_pdl_excel="",
        stato_attivita_excel="",
        data_controllo_excel="",
        personale_excel=""
    )

    with patch.object(automator, "_controlla_e_solleva_timeout") as mock_chk:
        def side_effect(selectors, timeout=None):
            if selectors == Config.ICON_GIA_PRENOTATO_SELECTORS:
                return MagicMock()
            if selectors == Config.ICON_DA_PRENOTARE_SELECTORS:
                raise AutomationError("Non da prenotare")
            return MagicMock()

        with patch.object(automator, "_wait_and_find", side_effect=side_effect):
            res = automator.prenota_singolo_pdl(pdl)
            assert res == "Già Prenotato"


@patch("src.automation.actions.WebDriverWait")
@patch("src.automation.actions.time.sleep")
def test_prenota_singolo_pdl_not_found(mock_sleep: MagicMock, mock_wait: MagicMock, automator: SafeWorkAutomator) -> None:
    """Testa il caso in cui il PDL inserito non esista."""
    mock_wait_instance = MagicMock()
    mock_wait.return_value = mock_wait_instance

    pdl = PDLData(4, "99999", "Sud", "U100", "Lavoro", "", "", "")

    def side_effect(selectors, timeout=None):
        if selectors == Config.MSG_PDL_NON_TROVATO_SELECTORS:
            return MagicMock()
        if selectors in [Config.ICON_GIA_PRENOTATO_SELECTORS, Config.ICON_DA_PRENOTARE_SELECTORS]:
            raise AutomationError("Errore")
        return MagicMock()

    with patch.object(automator, "_wait_and_find", side_effect=side_effect):
        res = automator.prenota_singolo_pdl(pdl)
        assert res == "PDL non trovato"


@patch("src.automation.actions.WebDriverWait")
@patch("src.automation.actions.time.sleep")
def test_prenota_singolo_pdl_success(mock_sleep: MagicMock, mock_wait: MagicMock, automator: SafeWorkAutomator) -> None:
    """Testa il flusso di success per la prenotazione di un PDL."""
    mock_wait_instance = MagicMock()
    mock_wait.return_value = mock_wait_instance

    pdl = PDLData(4, "12345", "Sud", "U100", "Lavoro", "", "", "")

    def side_effect(selectors, timeout=None):
        if selectors == Config.ICON_DA_PRENOTARE_SELECTORS:
            return MagicMock()
        if selectors in [Config.ICON_GIA_PRENOTATO_SELECTORS, Config.MSG_PDL_NON_TROVATO_SELECTORS]:
            raise AutomationError("Non presente")
        return MagicMock()

    with patch.object(automator, "_wait_and_find", side_effect=side_effect):
        with patch.object(automator, "_click_element") as mock_click:
            res = automator.prenota_singolo_pdl(pdl)
            assert res == "Prenotazione Eseguita"
            assert mock_click.call_count >= 1


@patch("src.automation.actions.WebDriverWait")
def test_handle_timeout_alert(mock_wait: MagicMock, automator: SafeWorkAutomator) -> None:
    """Verifica il click sul tasto OK per superare l'alert di timeout del sito."""
    mock_wait_instance = MagicMock()
    mock_wait.return_value = mock_wait_instance

    mock_ok_button = MagicMock()
    mock_wait_instance.until.return_value = mock_ok_button

    automator._handle_timeout_alert()

    mock_ok_button.click.assert_called_once()
