"""Azioni di automazione sul portale SafeWork ISAB."""

import time
from collections.abc import Sequence
from contextlib import suppress
from typing import Any

from loguru import logger
from selenium import webdriver
from selenium.common.exceptions import (
    NoSuchElementException,
    TimeoutException,
)
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait

from ..config import Config
from ..models import AutomationError, PDLData, TimeoutAlertError


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
        self.current_pdl_context: str | None = None
        logger.info(f"SafeWorkAutomator inizializzato (dry_run={dry_run})")

    def _attendi_caricamento_pagina(
        self, timeout_overlay: int = Config.SELENIUM_TIMEOUT_LONG, timeout_popup: int = 3
    ) -> None:
        """Attende la scomparsa di overlay e indicatori di caricamento sistema."""
        logger.debug("Verifica indicatori di caricamento...")
        self._attendi_overlay_gis(timeout_overlay)
        self._attendi_span_caricamento(timeout_overlay)
        self._gestisci_popup_imprevisti(timeout_popup)
        time.sleep(Config.PAUSE_GENERAL_SHORT)

    def _attendi_overlay_gis(self, timeout: int) -> None:
        """Attende la scomparsa dell'overlay GIS."""
        try:
            wait = WebDriverWait(self.driver, timeout, poll_frequency=Config.POLLING_FREQUENCY)
            wait.until(ec.invisibility_of_element_located((By.ID, "GISWaitOverlay")))
        except TimeoutException:
            logger.warning(f"Overlay 'GISWaitOverlay' ancora presente dopo {timeout}s.")

    def _attendi_span_caricamento(self, timeout: int) -> None:
        """Attende la scomparsa dello span di caricamento testuale."""
        xpath = "//span[contains(text(), 'Caricamento...')]"
        try:
            with suppress(TimeoutException):
                WebDriverWait(self.driver, 1, poll_frequency=Config.POLLING_FREQUENCY).until(
                    ec.visibility_of_element_located((By.XPATH, xpath))
                )
            WebDriverWait(self.driver, timeout, poll_frequency=Config.POLLING_FREQUENCY).until(
                ec.invisibility_of_element_located((By.XPATH, xpath))
            )
        except TimeoutException:
            logger.debug("Timeout attesa span 'Caricamento...' (proseguo)")

    def _gestisci_popup_imprevisti(self, timeout: int) -> None:
        """Tenta di rilevare e chiudere popup modali non attesi."""
        try:
            wait = WebDriverWait(self.driver, timeout)
            modal = wait.until(
                ec.visibility_of_element_located(
                    (By.XPATH, "//div[contains(@class, 'modal') and contains(@style, 'display: block')]")
                )
            )
            logger.warning("Rilevato popup modale imprevisto. Tento la chiusura.")
            if not self.dry_run:
                self._chiudi_modal(modal)
        except TimeoutException:
            pass

    def _chiudi_modal(self, modal: WebElement) -> None:
        """Tenta di cliccare il pulsante di chiusura di un modal."""
        try:
            btn = modal.find_element(
                By.XPATH,
                ".//*[self::button or self::span or self::a][contains(text(), 'OK') or contains(text(), 'Si') or contains(text(), 'Yes') or @data-dismiss='modal']",
            )
            btn.click()
        except NoSuchElementException:
            try:
                modal.find_element(By.CSS_SELECTOR, "*[idtxt='E421C594']").click()
            except NoSuchElementException:
                logger.error("Impossibile trovare pulsante di chiusura nel modale.")
        with suppress(Exception):
            WebDriverWait(self.driver, 10).until(ec.invisibility_of_element(modal))

    def _find_element(
        self,
        selectors: Sequence[tuple[By | str, str]],
        timeout: int | None = None,
        condition: Any = ec.visibility_of_element_located,
    ) -> Any:
        """Cerca un elemento provando diversi selettori in sequenza."""
        wait_time = timeout or Config.SELENIUM_TIMEOUT_MEDIUM
        wait = WebDriverWait(self.driver, wait_time, poll_frequency=Config.POLLING_FREQUENCY)

        for by, value in selectors:
            try:
                return wait.until(condition((by, value)))
            except (TimeoutException, Exception):
                continue

        pdl_info = f" [PDL: {self.current_pdl_context}]" if self.current_pdl_context else ""
        raise AutomationError(f"Impossibile trovare l'elemento con i selettori forniti{pdl_info}.")

    def _click_element(
        self, selectors: Sequence[tuple[By | str, str]], name: str, timeout: int | None = None
    ) -> None:
        """Esegue il click su un elemento con logiche di 'click robusto' (Overlay check + JS Fallback)."""
        logger.info(f"Click su '{name}'")
        if self.dry_run:
            logger.info(f"[DRY RUN] Salto click su '{name}'")
            return

        self._attendi_caricamento_pagina()
        el = self._find_element(selectors, timeout, ec.element_to_be_clickable)
        self._esegui_click_robusto(el, name)

    def _esegui_click_robusto(self, el: WebElement, name: str) -> None:
        """Tenta diverse strategie di click per massimizzare il successo."""
        try:
            el.click()
        except Exception:
            try:
                self.driver.execute_script("arguments[0].click();", el)
                logger.debug(f"Click su '{name}' eseguito via JavaScript.")
            except Exception:
                try:
                    ActionChains(self.driver).move_to_element(el).click().perform()
                    logger.debug(f"Click su '{name}' eseguito via ActionChains.")
                except Exception as e_final:
                    logger.error(f"Errore critico durante il click su '{name}': {e_final}")
                    raise AutomationError(f"Impossibile cliccare su '{name}': {e_final}") from e_final

    def _inserisci_testo(self, selectors: Sequence[tuple[By | str, str]], text: str, name: str) -> None:
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
            selectors = Config.TIMEOUT_ALERT_P_TAG_SELECTORS + Config.TIMEOUT_ALERT_FALLBACK_SELECTORS
            for by, val in selectors:
                self._controlla_e_solleva_timeout(by, val, context)
        except TimeoutAlertError:
            raise
        except Exception:
            pass

    def _controlla_e_solleva_timeout(self, by: str | By, val: str, context: str) -> None:
        """Controlla se un elemento di timeout è visualizzato e solleva l'eccezione."""
        for el in self.driver.find_elements(by, val):
            if el.is_displayed():
                msg = el.text
                logger.error(f"Alert timeout rilevato in '{context}': {msg}")
                self._handle_timeout_alert()
                raise TimeoutAlertError(f"Timeout sito: {msg}")

    def _handle_timeout_alert(self) -> None:
        """Tenta di chiudere l'alert di timeout cliccando su OK."""
        try:
            btn = self._find_element(Config.TIMEOUT_ALERT_OK_BUTTON_SELECTORS, 5, ec.element_to_be_clickable)
            btn.click()
            time.sleep(Config.PAUSE_GENERAL_MEDIUM)
        except Exception:
            logger.warning("Impossibile chiudere l'alert di timeout.")

    def login(self, url: str, user: str, password: str) -> None:
        """Esegue la procedura di login al portale."""
        logger.info(f"Accesso a {url} per l'utente {user}")
        self.driver.get(url)
        time.sleep(Config.PAUSA_DOPO_GET_INIZIALE)

        self._click_element(Config.DROPDOWN_SITO_SELECTORS, "Dropdown Selezione Sito")
        time.sleep(Config.PAUSA_DOPO_CLICK_DROPDOWN_SITO)
        self._click_element(Config.ISAB_SUD_OPTION_SELECTORS, "Opzione ISAB Sud")
        time.sleep(Config.PAUSA_DOPO_SELEZIONE_SITO)

        self._inserisci_testo(Config.USERNAME_FIELD_SELECTORS, user, "Username")
        self._inserisci_testo(Config.PASSWORD_FIELD_SELECTORS, password, "Password")
        self._click_element(Config.LOGIN_BUTTON_SELECTORS, "Pulsante Login")

        wait = WebDriverWait(self.driver, Config.SELENIUM_TIMEOUT_LONG)
        wait.until(ec.invisibility_of_element_located(Config.USERNAME_FIELD_SELECTORS[0]))

        self._attendi_caricamento_pagina()
        self.check_for_timeout_alert("Login")
        logger.info("Login completato con successo.")

    def navigate_to_booking(self) -> None:
        """Naviga alla sezione Prenotazione PDL."""
        self._attendi_caricamento_pagina()
        self._click_element(Config.HOME_BUTTON_SELECTORS, "Pulsante Home")
        self._attendi_caricamento_pagina()

        link = self._find_element(
            Config.LINK_PRENOTAZIONE_PDL_SELECTORS, condition=ec.element_to_be_clickable
        )
        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", link)
        time.sleep(0.5)
        link.click()

        self._attendi_caricamento_pagina()
        self.check_for_timeout_alert("Navigazione Prenotazione")
        logger.info("Pagina Prenotazione PDL raggiunta.")

    def process_pdl(self, data: PDLData) -> str:
        """Esegue la ricerca e l'eventuale prenotazione di un singolo PDL."""
        self.current_pdl_context = data.pdl
        logger.info(f"Elaborazione PDL {data.pdl}...")

        self._inserisci_testo(Config.INPUT_PDL_WEB_SELECTORS, data.pdl, "Input Ricerca PDL")
        self._click_element(Config.TASTO_CERCA_PDL_SELECTORS, "Pulsante Ricerca")
        self._attendi_caricamento_pagina()
        self.check_for_timeout_alert(f"Ricerca PDL {data.pdl}")

        if self._is_element_present(Config.MSG_PDL_NON_TROVATO_SELECTORS):
            logger.info(f"PDL {data.pdl} non trovato nel sistema.")
            return "Non Trovato"

        if self._is_element_present(Config.ICON_GIA_PRENOTATO_SELECTORS):
            logger.info(f"PDL {data.pdl} risulta già prenotato.")
            return "Già Prenotato"

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

    def _is_element_present(self, selectors: Sequence[tuple[By | str, str]], timeout: int = 5) -> bool:
        """Verifica rapida della presenza di un elemento."""
        try:
            WebDriverWait(self.driver, timeout)
            return any(self.driver.find_elements(by, val) for by, val in selectors)
        except Exception:
            return False
