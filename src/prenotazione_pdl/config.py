"""Configurazioni e costanti per l'automazione SafeWork."""

import os
from selenium.webdriver.common.by import By

# Determinazione della directory dello script
SCRIPT_FILE_PATH_ABS = os.path.abspath(__file__)
# Saliamo di tre livelli: src/prenotazione_pdl/config.py -> src/prenotazione_pdl/ -> src/ -> root
BASE_DIRECTORY = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

class Config:
    """Classe contenente tutte le costanti di configurazione del processo."""
    
    SCRIPT_DIR = BASE_DIRECTORY
    FILE_STATO_PROCESSO = "stato_processo_pdl.json"
    DEFAULT_URL_SITO = "https://safework.isab.com/"
    EXCEL_FILE_CONFIG_NAME = "parametri prenotazione pdl.xlsx"
    
    NOME_FOGLIO_DATI_PDL = "Riepilogo"
    EXCEL_SHEET_CREDENTIALS = "credenziali"
    EXCEL_SHEET_PERCORSI = "percorsi"
    
    CELLA_URL_SITO = "B3"
    CELLA_PERCORSO_FILE_DATI_PDL = "B2"
    USERNAME_CELL_EXCEL = "A3"
    PASSWORD_CELL_EXCEL = "B3"
    
    COLONNE_EXCEL_REPORT = {
        'pdl': 'E', 'area': 'D', 'descrizione': 'G', 'stato_pdl_excel': 'M',
        'stato_attivita_excel': 'Q', 'data_controllo_excel': 'R', 'personale_excel': 'S',
        'impianto': 'F'
    }
    
    RIGA_INIZIO_DATI_PDL = 4
    MAPPA_GIORNI_COLONNE_DATE = {i: chr(ord('H') + i) for i in range(5)}
    
    # Macro Excel
    MACRO_SEQ_1 = "ordina.OrdinaEFormattaTabellaCorrente"
    MACRO_SEQ_2 = "Modulo8.aggiornaQuery"
    MACRO_SEQ_3 = "Rimuovi_Tutti_I_Filtri.RimuoviTuttiIFiltri"
    
    # Selettori Selenium
    DROPDOWN_SITO_SELECTORS = [(By.XPATH, "//button[@class='ms-choice']")]
    ISAB_SUD_OPTION_SELECTORS = [(By.XPATH, "//div[contains(@class, 'ms-drop')]//label[.//span[normalize-space()='ISAB Sud']]")]
    USERNAME_FIELD_SELECTORS = [(By.ID, "inpUtente")]
    PASSWORD_FIELD_SELECTORS = [(By.ID, "inpPassword")]
    LOGIN_BUTTON_SELECTORS = [(By.ID, "btnLogin")]
    INDICATORE_CARICAMENTO_LOGIN_SELECTORS = [(By.XPATH, "//*[contains(text(), 'Caricamento...')]")]
    
    HOME_BUTTON_SELECTORS = [(By.ID, "topIcon-actHomePage")]
    LINK_PRENOTAZIONE_PDL_SELECTORS = [(By.ID, "sideBar-actPrenotazionePdL")]
    INPUT_PDL_WEB_SELECTORS = [(By.ID, "inpNumPermessoApparecchitura")]
    TASTO_CERCA_PDL_SELECTORS = [(By.ID, "btnAvviaRicerca")]
    
    ICON_DA_PRENOTARE_SELECTORS = [(By.XPATH, "//i[@title='Da prenotare' and contains(@class, 'daprenotare')]")]
    ICON_GIA_PRENOTATO_SELECTORS = [(By.XPATH, "//i[@title='Prenotato' and contains(@class, 'prenotato')]")]
    TASTO_SALVA_PRENOTAZIONE_SELECTORS = [(By.ID, "btnSalva")]
    
    MSG_PDL_NON_TROVATO_SELECTORS = [(By.XPATH, "//span[normalize-space(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZÀÈÉÌÒÙ', 'abcdefghijklmnopqrstuvwxyzàèéìòù'))='nessun permesso trovato']")]
    
    TIMEOUT_ALERT_P_TAG_SELECTORS = [(By.XPATH, "//p[@idtxt='CD69E8EA'][contains(text(), 'Request timeout (30000ms)')]")]
    TIMEOUT_ALERT_FALLBACK_SELECTORS = [(By.XPATH, "//div[contains(@class, 'bootstrap-dialog-message')]//p[contains(text(), 'Request timeout (30000ms)')]")]
    TIMEOUT_ALERT_OK_BUTTON_SELECTORS = [(By.XPATH, "//div[contains(@class, 'bootstrap-dialog') or contains(@class, 'modal-dialog')]//button[normalize-space(translate(., 'OKok', 'okok'))='ok']")]
    
    # Timeout e Tentativi
    RETRY_GENERAL_ATTEMPTS = 3
    RETRY_WAIT_MULTIPLIER = 1
    RETRY_WAIT_MIN_SECONDS = 2
    RETRY_WAIT_MAX_SECONDS = 10
    MAX_SETUP_ATTEMPTS = 5
    
    SELENIUM_TIMEOUT_PAGE_LOAD = 90
    SELENIUM_TIMEOUT_LONG = 120
    SELENIUM_TIMEOUT_MEDIUM = 40
    SELENIUM_TIMEOUT_SHORT = 20
    SELENIUM_TIMEOUT_QUICK_CHECK = 5
    
    PAUSE_GENERAL_SHORT = 0.5
    PAUSE_GENERAL_MEDIUM = 1.5
    PAUSA_DOPO_GET_INIZIALE = 2
    PAUSA_DOPO_CLICK_DROPDOWN_SITO = 1.5
    PAUSA_DOPO_SELEZIONE_SITO = 1
    PAUSA_DOPO_CLICK_DA_PRENOTARE = 1.0
    PAUSA_TRA_PDL = 1.0
    PAUSA_TRA_TENTATIVI_SETUP_FALLITI = 3
    PAUSA_APERTURA_EXCEL = 3
    PAUSA_ATTIVAZIONE_FOGLIO_EXCEL = 1
    PAUSA_COMPLETAMENTO_MACRO = 10
