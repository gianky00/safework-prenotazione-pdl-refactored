"""Modulo per l'elaborazione dei file Excel e l'esecuzione di macro VBA."""

import logging
import os
import time
from datetime import date, datetime
from typing import Any, List, Optional, Tuple

import openpyxl
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type

from ..config import Config
from ..models import CriticalConfigError, PDLData

# Gestione opzionale di pywin32
WIN32COM_AVAILABLE = False
try:
    import win32com.client
    WIN32COM_AVAILABLE = True
except ImportError:
    pass

logger = logging.getLogger("PDLAutomator.Excel")

class ExcelProcessor:
    """Gestisce la lettura dei parametri da Excel e l'estrazione dei dati PDL."""

    def __init__(self, config_file_path: str, data_file_path: Optional[str] = None) -> None:
        """
        Inizializza il processore Excel.
        
        Args:
            config_file_path: Percorso del file Excel di configurazione principale.
            data_file_path: Percorso del file dati PDL (se già noto).
        """
        self.config_file_path = config_file_path
        self.data_file_path = data_file_path
        logger.info(f"ExcelProcessor inizializzato. Config: '{config_file_path}'")
        if not os.path.exists(self.config_file_path):
            raise CriticalConfigError(f"File parametri Excel non trovato: {self.config_file_path}")

    @retry(
        stop=stop_after_attempt(Config.RETRY_GENERAL_ATTEMPTS),
        wait=wait_exponential(multiplier=Config.RETRY_WAIT_MULTIPLIER, min=Config.RETRY_WAIT_MIN_SECONDS, max=Config.RETRY_WAIT_MAX_SECONDS),
        retry=retry_if_exception_type(Exception),
        reraise=True
    )
    def _leggi_cella(self, file_path: str, sheet_name: str, cell_address: str, data_only: bool = True) -> Any:
        """Legge il valore di una singola cella con logica di retry."""
        try:
            wb = openpyxl.load_workbook(file_path, data_only=data_only)
            sheet = wb[sheet_name]
            val = sheet[cell_address].value
            wb.close()
            return str(val).strip() if isinstance(val, str) else val
        except Exception as e:
            logger.error(f"Errore lettura cella {cell_address} in {file_path}: {e}")
            raise

    def get_website_url(self) -> str:
        """Recupera l'URL del sito SafeWork dal file di configurazione."""
        url = self._leggi_cella(self.config_file_path, Config.EXCEL_SHEET_PERCORSI, Config.CELLA_URL_SITO)
        if not isinstance(url, str) or not url.startswith("http"):
            logger.warning(f"URL non valido in Excel ('{url}'). Uso fallback: {Config.DEFAULT_URL_SITO}")
            return Config.DEFAULT_URL_SITO
        return url

    def get_pdl_data_file_path(self) -> str:
        """Recupera il percorso del file dati PDL dal file di configurazione."""
        path = self._leggi_cella(self.config_file_path, Config.EXCEL_SHEET_PERCORSI, Config.CELLA_PERCORSO_FILE_DATI_PDL)
        if not isinstance(path, str):
            raise CriticalConfigError("Percorso file dati PdL non trovato o non valido in Excel.")
        
        self.data_file_path = os.path.normpath(path)
        if not os.path.isabs(self.data_file_path):
            # Se è relativo, lo consideriamo relativo alla cartella dello script
            self.data_file_path = os.path.join(Config.SCRIPT_DIR, self.data_file_path)
            
        logger.info(f"Percorso file dati PdL risolto: {self.data_file_path}")
        return self.data_file_path

    def get_credentials(self, interactive_pwd: bool = False) -> Tuple[str, str]:
        """Recupera username e password."""
        import getpass
        username = self._leggi_cella(self.config_file_path, Config.EXCEL_SHEET_CREDENTIALS, Config.USERNAME_CELL_EXCEL)
        if not username:
            raise CriticalConfigError("Username non trovato nel file di configurazione.")
            
        if interactive_pwd:
            password = getpass.getpass(f"Inserire password per '{username}': ")
        else:
            password = self._leggi_cella(self.config_file_path, Config.EXCEL_SHEET_CREDENTIALS, Config.PASSWORD_CELL_EXCEL)
            
        if not password:
            raise CriticalConfigError(f"Password non trovata per l'utente '{username}'.")
            
        return str(username), str(password)

    def run_pdl_macros(self) -> bool:
        """Esegue la sequenza di macro VBA necessaria per aggiornare i dati."""
        if not WIN32COM_AVAILABLE:
            logger.warning("Libreria pywin32 non disponibile. Esecuzione macro saltata.")
            return False
            
        if not self.data_file_path:
            self.get_pdl_data_file_path()
            
        if not self.data_file_path or not os.path.exists(self.data_file_path):
            logger.error(f"Impossibile eseguire macro: file '{self.data_file_path}' non trovato.")
            return False
            
        macros = [
            (Config.MACRO_SEQ_1, Config.NOME_FOGLIO_DATI_PDL),
            (Config.MACRO_SEQ_2, Config.NOME_FOGLIO_DATI_PDL),
            (Config.MACRO_SEQ_3, Config.NOME_FOGLIO_DATI_PDL)
        ]
        
        for macro_name, sheet_name in macros:
            if not self._esegui_macro_vba(macro_name, sheet_name):
                return False
        return True

    def _esegui_macro_vba(self, macro_name: str, sheet_to_activate: Optional[str] = None) -> bool:
        """Esegue una singola macro VBA utilizzando COM interop."""
        if not self.data_file_path: return False
        
        abs_path = os.path.abspath(self.data_file_path)
        logger.info(f"Esecuzione macro '{macro_name}' su '{abs_path}'")
        
        excel_app = None
        workbook = None
        success = False
        
        try:
            excel_app = win32com.client.Dispatch("Excel.Application")
            excel_app.Visible = False
            excel_app.DisplayAlerts = False
            
            workbook = excel_app.Workbooks.Open(abs_path)
            time.sleep(Config.PAUSA_APERTURA_EXCEL)
            
            if sheet_to_activate:
                try:
                    workbook.Sheets(sheet_to_activate).Activate()
                except Exception as e:
                    logger.warning(f"Impossibile attivare il foglio '{sheet_to_activate}': {e}")
            
            excel_app.Application.Run(macro_name)
            time.sleep(Config.PAUSA_COMPLETAMENTO_MACRO)
            
            workbook.Save()
            success = True
            logger.info(f"Macro '{macro_name}' completata con successo.")
        except Exception as e:
            logger.error(f"Errore durante l'esecuzione della macro '{macro_name}': {e}", exc_info=True)
            raise
        finally:
            if workbook:
                workbook.Close(SaveChanges=success)
            if excel_app:
                excel_app.Quit()
        return success

    def get_pdl_list_from_excel(self) -> List[PDLData]:
        """Estrae la lista dei PDL da processare in base al giorno corrente."""
        if not self.data_file_path or not os.path.exists(self.data_file_path):
            return []
            
        lista_pdl: List[PDLData] = []
        oggi = date.today()
        giorno_settimana = oggi.weekday()
        
        if giorno_settimana not in Config.MAPPA_GIORNI_COLONNE_DATE:
            logger.warning(f"Oggi ({oggi.strftime('%A')}) non è un giorno previsto per l'automazione.")
            return []
            
        colonna_giorno = Config.MAPPA_GIORNI_COLONNE_DATE[giorno_settimana]
        logger.info(f"Estrazione PDL per {oggi.isoformat()} (Colonna Excel: {colonna_giorno})")
        
        try:
            wb = openpyxl.load_workbook(self.data_file_path, data_only=True)
            sheet = wb[Config.NOME_FOGLIO_DATI_PDL]
            
            for riga in range(Config.RIGA_INIZIO_DATI_PDL, sheet.max_row + 1):
                cella_giorno = sheet[f"{colonna_giorno}{riga}"].value
                if str(cella_giorno).strip().upper() == 'X':
                    pdl_info = self._estrai_riga_pdl(sheet, riga)
                    if pdl_info.pdl:
                        lista_pdl.append(pdl_info)
            wb.close()
        except Exception as e:
            logger.error(f"Errore durante la lettura della lista PDL: {e}", exc_info=True)
            
        return lista_pdl

    def _estrai_riga_pdl(self, sheet: Any, riga: int) -> PDLData:
        """Helper per estrarre i dati di un PDL da una riga Excel."""
        data = {"riga_excel_debug": riga}
        for key, col in Config.COLONNE_EXCEL_REPORT.items():
            val = sheet[f"{col}{riga}"].value
            if isinstance(val, datetime):
                data[key] = val.strftime("%d/%m/%Y")
            elif isinstance(val, (int, float)) and key == 'pdl':
                data[key] = str(int(val))
            elif val is not None:
                data[key] = str(val).strip()
            else:
                data[key] = ""
        return PDLData(**data) # type: ignore
