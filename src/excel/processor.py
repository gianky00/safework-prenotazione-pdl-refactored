"""Modulo per l'elaborazione dei file Excel di configurazione e dati PDL."""

import contextlib
import getpass
import os
import subprocess
import time
from datetime import datetime
from typing import Any

import openpyxl
from loguru import logger
from openpyxl.utils import column_index_from_string
from tenacity import (
    retry,
    retry_if_exception_type,
    stop_after_attempt,
    wait_exponential,
)

from ..config import Config
from ..models import CriticalConfigError, PDLData

# Gestione opzionale win32com per macro Excel (solo Windows)
WIN32COM_AVAILABLE = False
try:
    import win32com.client

    WIN32COM_AVAILABLE = True
except ImportError:
    pass


class ExcelProcessor:
    """Gestisce la lettura dei parametri da Excel e l'estrazione dei dati PDL."""

    def __init__(self, config_file_path: str, data_file_path: str | None = None, prenotazione_oggi_per_oggi: bool = False) -> None:
        """
        Inizializza il processore Excel.

        Args:
            config_file_path: Percorso del file Excel di configurazione principale.
            data_file_path: Percorso del file dati PDL (se già noto).
            prenotazione_oggi_per_oggi: Se True, configura B6="NO" (OGGI PER OGGI). 
                                        Se False (default), configura B6="SI" (OGGI PER DOMANI).
        """
        self.config_file_path = config_file_path
        self.data_file_path = data_file_path
        self.prenotazione_oggi_per_oggi = prenotazione_oggi_per_oggi
        self._cached_pdl_list: list[PDLData] = []
        logger.info(f"ExcelProcessor inizializzato. Config: '{config_file_path}'")
        if not os.path.exists(self.config_file_path):
            raise CriticalConfigError(f"File parametri Excel non trovato: {self.config_file_path}")

    @retry(
        stop=stop_after_attempt(Config.RETRY_GENERAL_ATTEMPTS),
        wait=wait_exponential(
            multiplier=Config.RETRY_WAIT_MULTIPLIER, min=Config.RETRY_WAIT_MIN_SECONDS, max=Config.RETRY_WAIT_MAX_SECONDS
        ),
        retry=retry_if_exception_type(Exception),
        reraise=True,
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

    def get_credentials(self, interactive_pwd: bool = False) -> tuple[str, str]:
        """Recupera username e password."""
        username = self._leggi_cella(self.config_file_path, Config.EXCEL_SHEET_CREDENTIALS, Config.USERNAME_CELL_EXCEL)
        if not username:
            raise CriticalConfigError("Username non trovato nel file di configurazione.")

        if interactive_pwd:
            password = getpass.getpass(f"Inserire password per '{username}': ")
        else:
            password = self._leggi_cella(
                self.config_file_path, Config.EXCEL_SHEET_CREDENTIALS, Config.PASSWORD_CELL_EXCEL
            )

        if not password:
            raise CriticalConfigError(f"Password non trovata per l'utente '{username}'.")

        return str(username), str(password)

    def run_pdl_macros(self) -> bool:
        """Esegue la sequenza di macro VBA necessaria per aggiornare i dati ed estrae i PDL."""
        if not WIN32COM_AVAILABLE:
            logger.warning("Libreria pywin32 non disponibile. Esecuzione macro saltata.")
            return False

        if not self.data_file_path:
            self.get_pdl_data_file_path()

        if not self.data_file_path or not os.path.exists(self.data_file_path):
            logger.error(f"Impossibile eseguire macro: file '{self.data_file_path}' non trovato.")
            return False

        self._cached_pdl_list = self._esegui_sessione_macro(os.path.abspath(self.data_file_path), run_updates=True)
        return len(self._cached_pdl_list) > 0

    def _chiudi_excel_forzatamente(self) -> None:
        """Tenta di chiudere eventuali processi Excel pendenti per evitare lock sui file."""
        try:
            logger.info("Tentativo di chiusura forzata di eventuali processi Excel attivi...")
            subprocess.run(["taskkill", "/F", "/IM", "excel.exe", "/T"], capture_output=True, check=False)
            time.sleep(2)
        except Exception as e:
            logger.warning(f"Non è stato possibile chiudere Excel forzatamente: {e}")

    def _esegui_sessione_macro(self, abs_path: str, run_updates: bool = True) -> list[PDLData]:
        """Gestisce una sessione Excel per l'esecuzione delle macro e l'estrazione dati."""
        self._chiudi_excel_forzatamente()
        excel_app, workbook = None, None
        extracted_data: list[PDLData] = []
        try:
            excel_app = win32com.client.Dispatch("Excel.Application")
            excel_app.Visible, excel_app.DisplayAlerts = False, False
            excel_app.AskToUpdateLinks = False

            logger.info(f"Apertura workbook: {abs_path}")
            workbook = excel_app.Workbooks.Open(abs_path, UpdateLinks=False, ReadOnly=(not run_updates))
            time.sleep(Config.PAUSA_APERTURA_EXCEL)

            # --- LOGICA PRENOTAZIONE "OGGI PER DOMANI" vs "OGGI PER OGGI" ---
            # Se prenotazione_oggi_per_oggi è True -> "OGGI PER OGGI" (B6="NO")
            # Se prenotazione_oggi_per_oggi è False -> "OGGI PER DOMANI" (B6="SI")
            valore_flag = "NO" if self.prenotazione_oggi_per_oggi else "SI"
            modalita_str = "OGGI PER OGGI" if self.prenotazione_oggi_per_oggi else "OGGI PER DOMANI"

            logger.info(f"Configurazione modalità '{modalita_str}': imposto B6 su '{valore_flag}'")
            try:
                sheet_ins = workbook.Sheets(Config.EXCEL_SHEET_INSERIMENTO_DATI)
                sheet_ins.Range(Config.CELLA_PRENOTAZIONE_OGGI).Value = valore_flag
            except Exception as e:
                logger.warning(f"Impossibile impostare il flag {valore_flag} in B6: {e}")

            if run_updates:
                for macro_name in [Config.MACRO_SEQ_1, Config.MACRO_SEQ_2, Config.MACRO_SEQ_3]:
                    logger.info(f"Esecuzione macro di aggiornamento '{macro_name}'...")
                    with contextlib.suppress(Exception):
                        excel_app.Application.Run(macro_name)
                    time.sleep(1)

            logger.info(f"Esecuzione macro di filtraggio '{Config.MACRO_FILTRO}'...")
            try:
                excel_app.Application.Run(Config.MACRO_FILTRO)
            except Exception as e:
                logger.warning(f"Errore durante l'esecuzione della macro di filtraggio: {e}")

            time.sleep(2)
            extracted_data = self._estrai_dati_visibili_win32(workbook)

            if run_updates:
                logger.info(f"Esecuzione reset finale filtri '{Config.MACRO_SEQ_3}'...")
                with contextlib.suppress(Exception):
                    excel_app.Application.Run(Config.MACRO_SEQ_3)

                # Ripristino flag prenotazione "OGGI" su "NO" per debug/consistenza futura
                logger.info("Ripristino modalità standard: imposto B6 su 'NO'")
                try:
                    sheet_ins = workbook.Sheets(Config.EXCEL_SHEET_INSERIMENTO_DATI)
                    sheet_ins.Range(Config.CELLA_PRENOTAZIONE_OGGI).Value = "NO"
                except Exception as e:
                    logger.warning(f"Impossibile ripristinare il flag B6 su 'NO': {e}")

                workbook.Save()

            logger.info(f"Sessione Excel completata. Estratti {len(extracted_data)} PDL.")

        except Exception as e:
            logger.error(f"Errore critico durante la sessione Excel: {e}", exc_info=True)
            if not run_updates: # Se fallisce in read-only, riproviamo con openpyxl come fallback estremo
                return []
            raise
        finally:
            if workbook:
                workbook.Close(SaveChanges=run_updates)
            if excel_app:
                excel_app.Quit()
        return extracted_data

    def _estrai_dati_visibili_win32(self, workbook: Any) -> list[PDLData]:
        """Estrae i dati dalle righe visibili dopo il filtraggio usando win32com."""
        xl_cell_type_visible = 12
        lista_pdl: list[PDLData] = []
        try:
            sheet = workbook.Sheets(Config.NOME_FOGLIO_DATI_PDL)
            last_row = sheet.UsedRange.Rows.Count + sheet.UsedRange.Row - 1
            if last_row < Config.RIGA_INIZIO_DATI_PDL:
                return []

            # Leggiamo tutto il range utile in una matrice per velocità (fino alla colonna S)
            max_col = max(column_index_from_string(c) for c in Config.COLONNE_EXCEL_REPORT.values())
            raw_range = sheet.Range(sheet.Cells(1, 1), sheet.Cells(last_row, max_col))
            matrix = raw_range.Value

            # Identifichiamo le righe visibili
            data_range = sheet.Range(sheet.Cells(Config.RIGA_INIZIO_DATI_PDL, 1), sheet.Cells(last_row, 1))
            try:
                visible_cells = data_range.SpecialCells(xl_cell_type_visible)

            except Exception:
                logger.warning("Nessuna riga visibile trovata dopo il filtraggio.")
                return []

            col_map = self._genera_col_map()

            for area in visible_cells.Areas:
                for r in range(area.Row, area.Row + area.Rows.Count):
                    riga_dati = matrix[r - 1]
                    pdl_data = self._crea_pdl_data_da_matrice(riga_dati, r, col_map)
                    if pdl_data and pdl_data.pdl:
                        lista_pdl.append(pdl_data)
        except Exception as e:
            logger.error(f"Errore durante l'estrazione win32: {e}")
        return lista_pdl

    def _crea_pdl_data_da_matrice(self, riga_dati: tuple[Any, ...], riga_idx: int, col_map: dict[str, int]) -> PDLData | None:
        """Converte una riga della matrice win32 in un oggetto PDLData."""
        data: dict[str, Any] = {"riga_excel_debug": riga_idx}
        for key, col_idx in col_map.items():
            val = riga_dati[col_idx - 1]
            data[key] = self._formatta_valore_cella(val, key)
        return PDLData(**data)

    def get_pdl_list_from_excel(self) -> list[PDLData]:
        """Restituisce la lista dei PDL (usando la cache o estraendoli se necessario)."""
        if self._cached_pdl_list:
            return self._cached_pdl_list

        if not self.data_file_path:
            self.get_pdl_data_file_path()

        if not self.data_file_path or not os.path.exists(self.data_file_path):
            return []

        # Se la cache è vuota (es. dry run o errore precedente), proviamo un'estrazione veloce read-only
        logger.info("Cache PDL vuota. Esecuzione estrazione rapida (read-only)...")
        self._cached_pdl_list = self._esegui_sessione_macro(os.path.abspath(self.data_file_path), run_updates=False)
        return self._cached_pdl_list

    def _genera_col_map(self) -> dict[str, int]:
        """Genera la mappatura nome_colonna -> indice_colonna."""
        return {key: column_index_from_string(col) for key, col in Config.COLONNE_EXCEL_REPORT.items()}

    def _formatta_valore_cella(self, val: Any, key: str) -> str:
        """Uniforma il formato dei valori letti da Excel."""
        if isinstance(val, datetime):
            return val.strftime("%d/%m/%Y")
        if isinstance(val, (int, float)) and key == 'pdl':
            # Gestione PDL come "123456.0" -> "123456"
            return str(int(val)) if val == int(val) else str(val)
        return str(val).strip() if val is not None else ""
