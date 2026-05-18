"""Gestione del ciclo di vita del WebDriver Selenium."""

import time

from loguru import logger
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

from ..config import Config
from ..models import AutomationError


class WebDriverManager:
    """Gestisce la creazione, il riavvio e la chiusura del driver Chrome."""

    def __init__(self, headless: bool = False, start_maximized: bool = True) -> None:
        """
        Inizializza il manager del driver.

        Args:
            headless: Se True, il browser verrà eseguito senza interfaccia grafica.
            start_maximized: Se True, avvia il browser massimizzato.
        """
        self.headless = headless
        self.start_maximized = start_maximized
        self.driver: webdriver.Chrome | None = None
        logger.info(f"WebDriverManager inizializzato (headless={headless}, start_maximized={start_maximized})")

    def get_driver(self) -> webdriver.Chrome:
        """Restituisce il driver esistente o ne crea uno nuovo."""
        if self.driver and self._is_driver_alive():
            logger.debug("Restituisco driver esistente.")
            return self.driver

        logger.info("Creazione di un nuovo WebDriver Chrome.")
        options = self._crea_opzioni_chrome()

        try:
            self.driver = webdriver.Chrome(options=options)
            self.driver.set_page_load_timeout(Config.SELENIUM_TIMEOUT_PAGE_LOAD)
            logger.info("WebDriver Chrome creato con successo.")
        except Exception as e:
            logger.error(f"Errore critico nella creazione del WebDriver: {e}", exc_info=True)
            raise AutomationError(f"Impossibile creare il WebDriver: {e}") from e

        return self.driver

    def _crea_opzioni_chrome(self) -> Options:
        """Configura le opzioni e le preferenze per Chrome."""
        options = Options()
        if self.headless:
            options.add_argument("--headless=new")

        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--log-level=3")
        options.add_argument("--silent")
        options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])

        if self.start_maximized and not self.headless:
            options.add_argument("--start-maximized")
        else:
            options.add_argument("--window-size=1280,1024")

        options.add_experimental_option(
            "prefs",
            {
                "credentials_enable_service": False,
                "profile.password_manager_enabled": False,
                "profile.password_manager_leak_detection": False,
            },
        )
        return options

    def _is_driver_alive(self) -> bool:
        """Verifica se l'istanza del driver è ancora attiva."""
        if not self.driver:
            return False
        try:
            _ = self.driver.title
        except Exception:
            return False
        else:
            return True

    def quit_driver(self) -> None:
        """Chiude il driver e pulisce l'istanza."""
        if self.driver:
            logger.info("Chiusura del WebDriver Chrome.")
            try:
                self.driver.quit()
            except Exception:
                pass
            finally:
                self.driver = None

    def restart_driver(self) -> webdriver.Chrome:
        """Riavvia il driver chiudendo la vecchia istanza."""
        logger.info("Riavvio del WebDriver richiesto.")
        self.quit_driver()
        time.sleep(Config.PAUSE_GENERAL_MEDIUM)
        return self.get_driver()
