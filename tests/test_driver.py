"""Test unitari per il modulo automation/driver.py."""

from unittest.mock import MagicMock, patch
import pytest
from src.automation.driver import WebDriverManager
from src.models import AutomationError


@pytest.fixture
def mock_chrome() -> MagicMock:
    """Fixture per mockare selenium.webdriver.Chrome."""
    with patch("src.automation.driver.webdriver.Chrome") as mock_c:
        yield mock_c


@pytest.fixture
def mock_options() -> MagicMock:
    """Fixture per mockare selenium.webdriver.chrome.options.Options."""
    # In src/automation/driver.py viene importato come:
    # from selenium.webdriver.chrome.options import Options
    with patch("src.automation.driver.Options") as mock_opt:
        yield mock_opt


def test_driver_manager_init() -> None:
    """Verifica l'inizializzazione del manager con le varie opzioni."""
    manager = WebDriverManager(headless=True, start_maximized=False)
    assert manager.headless is True
    assert manager.start_maximized is False
    assert manager.driver is None


def test_quit_driver_none() -> None:
    """Verifica che quit_driver non generi errore se il driver non è inizializzato."""
    manager = WebDriverManager()
    manager.quit_driver()
    assert manager.driver is None


def test_get_driver_success(mock_chrome: MagicMock, mock_options: MagicMock) -> None:
    """Verifica la creazione corretta del WebDriver e delle opzioni Chrome."""
    mock_driver_instance = MagicMock()
    mock_chrome.return_value = mock_driver_instance

    mock_options_instance = MagicMock()
    mock_options.return_value = mock_options_instance

    # Usiamo headless=False per testare start_maximized
    manager = WebDriverManager(headless=False, start_maximized=True)
    driver = manager.get_driver()

    assert driver == mock_driver_instance
    assert manager.driver == mock_driver_instance

    # Verifica le chiamate per le opzioni
    mock_options_instance.add_argument.assert_any_call("--start-maximized")
    mock_options_instance.add_argument.assert_any_call("--disable-gpu")

    # Verifica l'istanziazione di Chrome con le opzioni
    mock_chrome.assert_called_once_with(options=mock_options_instance)


def test_get_driver_success_headless(mock_chrome: MagicMock, mock_options: MagicMock) -> None:
    """Verifica la creazione delle opzioni in modalità headless."""
    mock_driver_instance = MagicMock()
    mock_chrome.return_value = mock_driver_instance

    mock_options_instance = MagicMock()
    mock_options.return_value = mock_options_instance

    manager = WebDriverManager(headless=True, start_maximized=True)
    manager.get_driver()

    # Dovrebbe aggiungere --headless=new e NON --start-maximized
    mock_options_instance.add_argument.assert_any_call("--headless=new")
    
    # Verifichiamo che --start-maximized non sia mai stato chiamato
    calls = [call[0][0] for call in mock_options_instance.add_argument.call_args_list]
    assert "--start-maximized" not in calls


def test_quit_driver_active(mock_chrome: MagicMock) -> None:
    """Verifica la dismissione corretta e sicura di un driver attivo."""
    mock_driver_instance = MagicMock()
    mock_chrome.return_value = mock_driver_instance

    manager = WebDriverManager()
    # Mocking _crea_opzioni_chrome per evitare dipendenze da Options
    with patch.object(WebDriverManager, "_crea_opzioni_chrome"):
        manager.get_driver()
        assert manager.driver is not None

        manager.quit_driver()
        mock_driver_instance.quit.assert_called_once()
        assert manager.driver is None


@patch("src.automation.driver.time.sleep")
def test_get_driver_retry_failure(mock_sleep: MagicMock, mock_chrome: MagicMock) -> None:
    """Verifica che se la creazione del driver fallisce continuamente, venga sollevata un'eccezione critica."""
    mock_chrome.side_effect = Exception("Errore driver continuo")

    manager = WebDriverManager()
    with patch.object(WebDriverManager, "_crea_opzioni_chrome"):
        with pytest.raises(AutomationError, match="Impossibile creare il WebDriver"):
            manager.get_driver()
