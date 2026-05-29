"""Test unitari per il modulo automation/actions.py."""

from unittest.mock import MagicMock, patch
import pytest
from selenium.common.exceptions import TimeoutException
from src.config import Config
from src.automation.actions import SafeWorkAutomator
from src.models import PDLData


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


@patch("src.automation.actions.WebDriverWait")
def test_login_success(mock_wait: MagicMock, automator: SafeWorkAutomator) -> None:
    """Testa il login con successo."""
    mock_wait_instance = MagicMock()
    mock_wait.return_value = mock_wait_instance
    
    automator.login("https://test.com", "user", "pwd")
    automator.driver.get.assert_called_once_with("https://test.com")


@patch("src.automation.actions.WebDriverWait")
def test_navigate_to_booking(mock_wait: MagicMock, automator: SafeWorkAutomator) -> None:
    """Testa la navigazione alla pagina prenotazione."""
    mock_wait_instance = MagicMock()
    mock_wait.return_value = mock_wait_instance
    
    automator.navigate_to_booking()
    assert automator.driver.find_element.called or mock_wait_instance.until.called


@patch("src.automation.actions.time.sleep")
def test_process_pdl_success(mock_sleep: MagicMock, automator: SafeWorkAutomator) -> None:
    """Testa il processamento PdL con successo."""
    pdl = PDLData(1, "123", "Area", "Imp", "Desc", "", "", "")
    
    with patch.object(automator, "check_for_timeout_alert"), \
         patch.object(automator, "_inserisci_testo"), \
         patch.object(automator, "_click_element"), \
         patch.object(automator, "_attendi_caricamento_pagina"), \
         patch.object(automator, "_is_element_present", return_value=False):
        
        res = automator.process_pdl(pdl)
        assert res == "Prenotazione Eseguita"


@patch("src.automation.actions.time.sleep")
def test_process_pdl_already_booked(mock_sleep: MagicMock, automator: SafeWorkAutomator) -> None:
    """Testa PdL già prenotato."""
    pdl = PDLData(1, "123", "Area", "Imp", "Desc", "", "", "")
    
    def side_effect_present(selectors, timeout=None):
        if selectors == Config.ICON_GIA_PRENOTATO_SELECTORS:
            return True
        return False

    with patch.object(automator, "check_for_timeout_alert"), \
         patch.object(automator, "_inserisci_testo"), \
         patch.object(automator, "_click_element"), \
         patch.object(automator, "_attendi_caricamento_pagina"), \
         patch.object(automator, "_is_element_present", side_effect=side_effect_present):
        
        res = automator.process_pdl(pdl)
        assert res == "Già Prenotato"


@patch("src.automation.actions.WebDriverWait")
def test_check_for_timeout_alert(mock_wait: MagicMock, automator: SafeWorkAutomator) -> None:
    """Testa il check degli alert di timeout."""
    mock_wait_instance = MagicMock()
    mock_wait.return_value = mock_wait_instance
    
    # Simula nessun alert trovato (timeout)
    mock_wait_instance.until.side_effect = TimeoutException()
    automator.check_for_timeout_alert("test_context") # Non deve esplodere
