"""Azioni di automazione sul portale SafeWork ISAB."""

import logging
import time
from typing import List, Optional, Tuple

from selenium import webdriver
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    NoSuchElementException,
    TimeoutException,
)
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from ..config import Config
from ..models import AutomationException, PDLData, TimeoutAlertDetected

logger = logging.getLogger("PDLAutomator.Actions")

class SafeWorkAutomator:
    """Implementa le interazioni specifiche con il sito SafeWork."""

    def __init__(self, driver: webdriver.Chrome, dry_run: bool = False) -> None:
        """
        Inizializza l'automatore.
        
        Args:
            driver: Istanza del WebDriver Chrome.
            dry_run: Se True, non esegue azioni di scrittura/modifica sul sito.
        """
        self.driver = driver
        self.dry_run = dry_run
        self.current_pdl_context: Optional[str] = None
        logger.info(f"SafeWorkAutomator inizializzato (dry_run={dry_run})")

    def _attendi_caricamento_pagina(
        self, 
        timeout_overlay: int = Config.SELENIUM_TIMEOUT_LONG, 
        timeout_popup: int = 3
    ) -> None:
        """Attende la scomparsa di overlay di caricamento e gestisce popup imprevisti."""
        logger.debug("Verifica overlay di caricamento...")
        try:
            wait = WebDriverWait(self.driver, timeout_overlay)
            wait.until(EC.invisibility_of_element_located((By.ID, "GISWaitOverlay")))
        except TimeoutException:
            logger.warning(f"Overlay 'GISWaitOverlay' ancora presente dopo {timeout_overlay}s.")

        # Gestione popup modali imprevisti
        try:
            wait_popup = WebDriverWait(self.driver, timeout_popup)
            modal = wait_popup.until(EC.visibility_of_element_located(
                (By.XPATH, "//div[contains(@class, 'modal') and contains(@style, 'display: block')]")
            ))
            logger.warning("Rilevato popup modale imprevisto. Tento la chiusura.")
            if not self.dry_run:
                btn = modal.find_element(By.XPATH, ".//button[contains(text(), 'OK') or @data-dismiss='modal']")
                btn.click()
                WebDriverWait(self.driver, 10).until(EC.invisibility_of_element(modal))
        except (TimeoutException, NoSuchElementException):
            pass
        
        time.sleep(Config.PAUSE_GENERAL_SHORT)

    def _find_element(
        self, 
        selectors: List[Tuple[By, str]], 
        timeout: Optional[int] = None,
        condition: Any = EC.visibility_of_element_located
    ) -> Any:
        """Cerca un elemento provando diversi selettori in sequenza."""
        wait_time = timeout or Config.SELENIUM_TIMEOUT_MEDIUM
        wait = WebDriverWait(self.driver, wait_time)
        
        for by, value in selectors:
            try:
                return wait.until(condition((by, value)))
            except (TimeoutException, Exception):
                continue
                
        pdl_info = f" [PDL: {self.current_pdl_context}]" if self.current_pdl_context else ""
        raise AutomationException(f"Impossibile trovare l'elemento con i selettori forniti{pdl_info}.")

    def _click_element(
        self, 
        selectors: List[Tuple[By, str]], 
        name: str, 
        timeout: Optional[int] = None
    ) -> None:
        """Esegue il click su un elemento con logiche di fallback (JS, ActionChains)."""
        logger.info(f"Click su '{name}'")
        if self.dry_run:
            logger.info(f"[DRY RUN] Salto click su '{name}'")
            return
            
        el = self._find_element(selectors, timeout, EC.element_to_be_clickable)
        try:
            # Preferiamo il click via JavaScript per bypassare sovrapposizioni minori
            self.driver.execute_script("arguments[0].click();", el)
        except Exception:
            try:
                el.click()
            except ElementClickInterceptedException:
                ActionChains(self.driver).move_to_element(el).click().perform()

    def _inserisci_testo(self, selectors: List[Tuple[By, str]], text: str, name: str) -> None:
        """Pulisce un campo e inserisce il testo specificato."""
        logger.info(f"Inserimento testo in '{name}'")
        if self.dry_run:
            logger.info(f"[DRY RUN] Salto inserimento in '{name}'")
            return
            
        el = self._find_element(selectors)
        el.click()
        el.send_keys(Keys.CONTROL + "a")
        el.send_keys(Keys.BACK_SPACE)
        if text:
            el.send_keys(text)

    def check_for_timeout_alert(self, context: str) -> None:
        """Verifica se è apparso un alert di timeout dal sistema SafeWork."""
        try:
            # Controllo rapido per alert di timeout
            selectors = Config.TIMEOUT_ALERT_P_TAG_SELECTORS + Config.TIMEOUT_ALERT_FALLBACK_SELECTORS
            for by, val in selectors:
                elements = self.driver.find_elements(by, val)
                for el in elements:
                    if el.is_displayed():
                        msg = el.text
                        logger.error(f"Alert timeout rilevato in '{context}': {msg}")
                        self._handle_timeout_alert()
                        raise TimeoutAlertDetected(f"Timeout sito: {msg}")
        except TimeoutAlertDetected:
            raise
        except Exception:
            pass

    def _handle_timeout_alert(self) -> None:
        """Tenta di chiudere l'alert di timeout cliccando su OK."""
        try:
            btn = self._find_element(Config.TIMEOUT_ALERT_OK_BUTTON_SELECTORS, 5, EC.element_to_be_clickable)
            btn.click()
            time.sleep(Config.PAUSE_GENERAL_MEDIUM)
        except Exception:
            logger.warning("Impossibile chiudere l'alert di timeout.")

    def login(self, url: str, user: str, password: str) -> None:
        """Esegue la procedura di login al portale."""
        logger.info(f"Accesso a {url} per l'utente {user}")
        self.driver.get(url)
        time.sleep(Config.PAUSA_DOPO_GET_INIZIALE)
        
        # Selezione sito (ISAB Sud)
        self._click_element(Config.DROPDOWN_SITO_SELECTORS, "Dropdown Selezione Sito")
        time.sleep(Config.PAUSA_DOPO_CLICK_DROPDOWN_SITO)
        self._click_element(Config.ISAB_SUD_OPTION_SELECTORS, "Opzione ISAB Sud")
        time.sleep(Config.PAUSA_DOPO_SELEZIONE_SITO)
        
        # Inserimento credenziali
        self._inserisci_testo(Config.USERNAME_FIELD_SELECTORS, user, "Username")
        self._inserisci_testo(Config.PASSWORD_FIELD_SELECTORS, password, "Password")
        self._click_element(Config.LOGIN_BUTTON_SELECTORS, "Pulsante Login")
        
        # Attesa caricamento dashboard
        wait = WebDriverWait(self.driver, Config.SELENIUM_TIMEOUT_LONG)
        wait.until(EC.invisibility_of_element_located(Config.USERNAME_FIELD_SELECTORS[0]))
        
        self._attendi_caricamento_pagina()
        self.check_for_timeout_alert("Login")
        logger.info("Login completato con successo.")

    def navigate_to_booking(self) -> None:
        """Naviga alla sezione Prenotazione PDL."""
        self._attendi_caricamento_pagina()
        self._click_element(Config.HOME_BUTTON_SELECTORS, "Pulsante Home")
        self._attendi_caricamento_pagina()
        
        link = self._find_element(Config.LINK_PRENOTAZIONE_PDL_SELECTORS, condition=EC.element_to_be_clickable)
        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", link)
        time.sleep(0.5)
        link.click()
        
        self._attendi_caricamento_pagina()
        self.check_for_timeout_alert("Navigazione Prenotazione")
        logger.info("Pagina Prenotazione PDL raggiunta.")

    def process_pdl(self, data: PDLData) -> str:
        """
        Esegue la ricerca e l'eventuale prenotazione di un singolo PDL.
        
        Returns:
            str: Lo stato finale dell'attività per questo PDL.
        """
        self.current_pdl_context = data.pdl
        logger.info(f"Elaborazione PDL {data.pdl}...")
        
        self._inserisci_testo(Config.INPUT_PDL_WEB_SELECTORS, data.pdl, "Input Ricerca PDL")
        self._click_element(Config.TASTO_CERCA_PDL_SELECTORS, "Pulsante Ricerca")
        self._attendi_caricamento_pagina()
        self.check_for_timeout_alert(f"Ricerca PDL {data.pdl}")
        
        # Verifica se non trovato
        if self._is_element_present(Config.MSG_PDL_NON_TROVATO_SELECTORS):
            logger.info(f"PDL {data.pdl} non trovato nel sistema.")
            return "Non Trovato"
            
        # Verifica se già prenotato
        if self._is_element_present(Config.ICON_GIA_PRENOTATO_SELECTORS):
            logger.info(f"PDL {data.pdl} risulta già prenotato.")
            return "Già Prenotato"
            
        # Tentativo di prenotazione
        try:
            self._click_element(Config.ICON_DA_PRENOTARE_SELECTORS, "Icona Da Prenotare")
            time.sleep(Config.PAUSA_DOPO_CLICK_DA_PRENOTARE)
            self._click_element(Config.TASTO_SALVA_PRENOTAZIONE_SELECTORS, "Pulsante Salva Prenotazione")
            self._attendi_caricamento_pagina()
            logger.info(f"PDL {data.pdl} prenotato con successo.")
            return "Prenotazione Eseguita"
        except Exception as e:
            logger.error(f"Errore durante la prenotazione del PDL {data.pdl}: {e}")
            return f"Errore ({type(e).__name__})"
        finally:
            self.current_pdl_context = None

    def _is_element_present(self, selectors: List[Tuple[By, str]], timeout: int = 5) -> bool:
        """Verifica rapida della presenza di un elemento."""
        try:
            wait = WebDriverWait(self.driver, timeout)
            for by, val in selectors:
                if self.driver.find_elements(by, val):
                    return True
            return False
        except Exception:
            return False
