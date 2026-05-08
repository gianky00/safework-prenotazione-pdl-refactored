"""Gestione del ciclo di vita del WebDriver Selenium."""

import logging
import time
from typing import Optional
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from ..models import AutomationException
from ..config import Config

logger = logging.getLogger("PDLAutomator.Driver")

class WebDriverManager:
    """Gestisce la creazione, il riavvio e la chiusura del driver Chrome."""

    def __init__(self, headless: bool = True, start_maximized: bool = False) -> None:
        """
        Inizializza il manager del driver.
        
        Args:
            headless: Se True, il browser verrà eseguito senza interfaccia grafica.
            start_maximized: Se True, il browser verrà avviato a schermo intero (solo se non headless).
        """
        self.headless = headless
        self.start_maximized = start_maximized
        self.driver: Optional[webdriver.Chrome] = None
        logger.info(f"WebDriverManager inizializzato (headless={headless}, start_maximized={start_maximized})")

    def get_driver(self) -> webdriver.Chrome:
        """
        Restituisce il driver esistente o ne crea uno nuovo se necessario.
        
        Returns:
            webdriver.Chrome: L'istanza del driver attivo.
        """
        if self.driver and self._is_driver_alive():
            logger.debug("Restituisco driver esistente.")
            return self.driver
            
        logger.info("Creazione di un nuovo WebDriver Chrome.")
        options = Options()
        if self.headless:
            options.add_argument("--headless=new")
            
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--log-level=3")
        options.add_argument("--disable-features=PasswordLeakDetection")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("enable-automation")
        options.add_argument("--disable-infobars")
        options.add_argument("--disable-browser-side-navigation")
        
        if self.start_maximized and not self.headless:
            options.add_argument("--start-maximized")
        else:
            # Dimensione fissa per consistenza
            options.add_argument("--window-size=1280,1024")
        
        prefs = {
            "credentials_enable_service": False,
            "profile.password_manager_enabled": False,
            "profile.password_manager_leak_detection": False
        }
        options.add_experimental_option("prefs", prefs)
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        
        try:
            self.driver = webdriver.Chrome(options=options)
            self.driver.set_page_load_timeout(Config.SELENIUM_TIMEOUT_PAGE_LOAD)
            logger.info("WebDriver Chrome creato con successo.")
        except Exception as e:
            logger.error(f"Errore critico nella creazione del WebDriver: {e}", exc_info=True)
            raise AutomationException(f"Impossibile creare il WebDriver: {e}") from e
            
        return self.driver

    def _is_driver_alive(self) -> bool:
        """Verifica se il driver è ancora attivo e rispondente."""
        try:
            _ = self.driver.title # type: ignore
            return True
        except Exception:
            return False

    def quit_driver(self) -> None:
        """Chiude il driver e pulisce l'istanza."""
        if self.driver:
            try:
                logger.info("Chiusura del WebDriver.")
                self.driver.quit()
            except Exception as e:
                logger.warning(f"Errore durante la chiusura del WebDriver: {e}")
            finally:
                self.driver = None

    def restart_driver(self) -> webdriver.Chrome:
        """Riavvia il driver chiudendo la vecchia istanza."""
        logger.info("Riavvio del WebDriver richiesto.")
        self.quit_driver()
        time.sleep(Config.PAUSE_GENERAL_MEDIUM)
        return self.get_driver()
