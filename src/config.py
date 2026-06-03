"""Configurazioni e costanti per l'automazione SafeWork."""

import os
from typing import ClassVar

from selenium.webdriver.common.by import By

# Determinazione della directory dello script
SCRIPT_FILE_PATH_ABS = os.path.abspath(__file__)
# Saliamo di due livelli: src/config.py -> src/ -> root
BASE_DIRECTORY = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


class Config:
    """Classe contenente tutte le costanti di configurazione del processo."""

    HEADLESS: ClassVar[bool] = True
    SCRIPT_DIR: ClassVar[str] = BASE_DIRECTORY
    LOGS_DIR: ClassVar[str] = os.path.join(BASE_DIRECTORY, "data", "logs")
    FILE_STATO_PROCESSO: ClassVar[str] = "stato_processo_pdl.json"
    DEFAULT_URL_SITO: ClassVar[str] = "https://safework.isab.com/"
    EXCEL_FILE_CONFIG_NAME: ClassVar[str] = "parametri prenotazione pdl.xlsx"

    NOME_FOGLIO_DATI_PDL: ClassVar[str] = "Riepilogo"
    EXCEL_SHEET_CREDENTIALS: ClassVar[str] = "credenziali"
    EXCEL_SHEET_PERCORSI: ClassVar[str] = "percorsi"
    EXCEL_SHEET_INSERIMENTO_DATI: ClassVar[str] = "Inserimento dati"

    CELLA_URL_SITO: ClassVar[str] = "B3"
    CELLA_PERCORSO_FILE_DATI_PDL: ClassVar[str] = "B2"
    CELLA_PRENOTAZIONE_OGGI: ClassVar[str] = "B6"  # SI = OGGI PER DOMANI, NO = OGGI PER OGGI
    USERNAME_CELL_EXCEL: ClassVar[str] = "A3"
    PASSWORD_CELL_EXCEL: ClassVar[str] = "B3"

    COLONNE_EXCEL_REPORT: ClassVar[dict[str, str]] = {
        "pdl": "E",
        "area": "D",
        "descrizione": "G",
        "stato_pdl_excel": "M",
        "stato_attivita_excel": "Q",
        "data_controllo_excel": "R",
        "personale_excel": "S",
        "impianto": "F",
    }

    RIGA_INIZIO_DATI_PDL: ClassVar[int] = 4

    # Macro Excel
    MACRO_FILTRO: ClassVar[str] = "FiltraProgrGiornalieraRiepilogoDinamico"
    MACRO_SEQ_1: ClassVar[str] = "ordina.OrdinaEFormattaTabellaCorrente"
    MACRO_SEQ_2: ClassVar[str] = "Modulo8.aggiornaQuery"
    MACRO_SEQ_3: ClassVar[str] = "Rimuovi_Tutti_I_Filtri.RimuoviTuttiIFiltri"

    # Selettori Selenium
    DROPDOWN_SITO_SELECTORS: ClassVar[list[tuple[str, str]]] = [(By.XPATH, "//button[@class='ms-choice']")]
    ISAB_SUD_OPTION_SELECTORS: ClassVar[list[tuple[str, str]]] = [
        (By.XPATH, "//div[contains(@class, 'ms-drop')]//label[.//span[normalize-space()='ISAB Sud']]")
    ]
    USERNAME_FIELD_SELECTORS: ClassVar[list[tuple[str, str]]] = [(By.ID, "inpUtente")]
    PASSWORD_FIELD_SELECTORS: ClassVar[list[tuple[str, str]]] = [(By.ID, "inpPassword")]
    LOGIN_BUTTON_SELECTORS: ClassVar[list[tuple[str, str]]] = [(By.ID, "btnLogin")]
    INDICATORE_CARICAMENTO_LOGIN_SELECTORS: ClassVar[list[tuple[str, str]]] = [
        (By.XPATH, "//*[contains(text(), 'Caricamento...')]")
    ]

    HOME_BUTTON_SELECTORS: ClassVar[list[tuple[str, str]]] = [(By.ID, "topIcon-actHomePage")]
    LINK_PRENOTAZIONE_PDL_SELECTORS: ClassVar[list[tuple[str, str]]] = [(By.ID, "sideBar-actPrenotazionePdL")]
    INPUT_PDL_WEB_SELECTORS: ClassVar[list[tuple[str, str]]] = [(By.ID, "inpNumPermessoApparecchitura")]
    TASTO_CERCA_PDL_SELECTORS: ClassVar[list[tuple[str, str]]] = [(By.ID, "btnAvviaRicerca")]

    ICON_DA_PRENOTARE_SELECTORS: ClassVar[list[tuple[str, str]]] = [
        (By.XPATH, "//i[@title='Da prenotare' and contains(@class, 'daprenotare')]")
    ]
    ICON_GIA_PRENOTATO_SELECTORS: ClassVar[list[tuple[str, str]]] = [
        (By.XPATH, "//i[@title='Prenotato' and contains(@class, 'prenotato')]")
    ]
    TASTO_SALVA_PRENOTAZIONE_SELECTORS: ClassVar[list[tuple[str, str]]] = [(By.ID, "btnSalva")]

    # Lista PDL Prenotati e Tempi Rimanenti
    BTN_LISTA_PRENOTATI_SELECTORS: ClassVar[list[tuple[str, str]]] = [(By.ID, "divPrenotazione")]
    TABELLA_PRENOTAZIONI_ROWS_SELECTORS: ClassVar[list[tuple[str, str]]] = [(By.CSS_SELECTOR, ".tabulator-row:not(.tabulator-group)")]
    CELL_NUM_PERMESSO_SELECTOR: ClassVar[str] = ".tabulator-cell[tabulator-field='NumPermesso']"
    CELL_TEMPO_RIMANENTE_SELECTOR: ClassVar[str] = ".tabulator-cell[tabulator-field='TempoRimanente']"

    MSG_PDL_NON_TROVATO_SELECTORS: ClassVar[list[tuple[str, str]]] = [
        (
            By.XPATH,
            "//span[normalize-space(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZÀÈÉÌÒÙ', 'abcdefghijklmnopqrstuvwxyzàèéìòù'))='nessun permesso trovato']",
        )
    ]

    TIMEOUT_ALERT_P_TAG_SELECTORS: ClassVar[list[tuple[str, str]]] = [
        (By.XPATH, "//p[@idtxt='CD69E8EA'][contains(text(), 'Request timeout (30000ms)')]")
    ]
    TIMEOUT_ALERT_FALLBACK_SELECTORS: ClassVar[list[tuple[str, str]]] = [
        (
            By.XPATH,
            "//div[contains(@class, 'bootstrap-dialog-message')]//p[contains(text(), 'Request timeout (30000ms)')]",
        )
    ]
    TIMEOUT_ALERT_OK_BUTTON_SELECTORS: ClassVar[list[tuple[str, str]]] = [
        (
            By.XPATH,
            "//div[contains(@class, 'bootstrap-dialog') or contains(@class, 'modal-dialog')]//button[normalize-space(translate(., 'OKok', 'okok'))='ok']",
        )
    ]

    # Timeout e Tentativi
    RETRY_GENERAL_ATTEMPTS: ClassVar[int] = 3
    RETRY_WAIT_MULTIPLIER: ClassVar[int] = 1
    RETRY_WAIT_MIN_SECONDS: ClassVar[int] = 2
    RETRY_WAIT_MAX_SECONDS: ClassVar[int] = 10
    MAX_SETUP_ATTEMPTS: ClassVar[int] = 5

    SELENIUM_TIMEOUT_PAGE_LOAD: ClassVar[int] = 90
    SELENIUM_TIMEOUT_LONG: ClassVar[int] = 120
    SELENIUM_TIMEOUT_MEDIUM: ClassVar[int] = 30
    SELENIUM_TIMEOUT_SHORT: ClassVar[int] = 10
    SELENIUM_TIMEOUT_QUICK_CHECK: ClassVar[int] = 3

    POLLING_FREQUENCY: ClassVar[float] = 0.2  # Ridotto per reattività maggiore (default selenium è 0.5)

    PAUSE_GENERAL_SHORT: ClassVar[float] = 0.2
    PAUSE_GENERAL_MEDIUM: ClassVar[float] = 0.5
    PAUSA_DOPO_GET_INIZIALE: ClassVar[float] = 1
    PAUSA_DOPO_CLICK_DROPDOWN_SITO: ClassVar[float] = 0.3
    PAUSA_DOPO_SELEZIONE_SITO: ClassVar[float] = 0.3
    PAUSA_DOPO_CLICK_DA_PRENOTARE: ClassVar[float] = 0.5
    PAUSA_TRA_PDL: ClassVar[float] = 0.5
    PAUSA_TRA_TENTATIVI_SETUP_FALLITI: ClassVar[int] = 3
    PAUSA_APERTURA_EXCEL: ClassVar[int] = 3
    PAUSA_ATTIVAZIONE_FOGLIO_EXCEL: ClassVar[int] = 1
    PAUSA_COMPLETAMENTO_MACRO: ClassVar[int] = 10
