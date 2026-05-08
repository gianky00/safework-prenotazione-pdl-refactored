"""Entry point principale per l'automazione Prenotazione PDL."""

import argparse
import json
import logging
import os
import shutil
import time
import traceback
from datetime import datetime
from typing import List, Tuple

from .automation.actions import SafeWorkAutomator
from .automation.driver import WebDriverManager
from .config import Config
from .excel.processor import ExcelProcessor
from .models import (
    AutomationException,
    CriticalConfigError,
    PDLData,
    TimeoutAlertDetected,
)

# Configurazione logger root
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)-8s - %(name)-25s - %(funcName)-25s - %(message)s"
)
logger = logging.getLogger("PDLAutomator")

class StateManager:
    """Gestisce la persistenza dello stato dell'elaborazione in un file JSON."""

    def __init__(self, file_path: str) -> None:
        self.file_path = file_path

    def carica_stato(self) -> Tuple[int, List[PDLData]]:
        """Carica l'ultimo indice processato e i risultati parziali."""
        if not os.path.exists(self.file_path):
            return -1, []
            
        try:
            with open(self.file_path, 'r', encoding='utf-8') as f:
                stato = json.load(f)
            idx = stato.get("ultimo_indice_pdl_processato", -1)
            dati = [PDLData(**d) for d in stato.get("risultati_elaborazione", [])]
            return idx, dati
        except Exception as e:
            logger.warning(f"File di stato corrotto o illeggibile: {e}")
            return -1, []

    def salva_stato(self, ultimo_indice: int, risultati: List[PDLData]) -> None:
        """Salva lo stato corrente su disco."""
        stato = {
            "ultimo_indice_pdl_processato": ultimo_indice,
            "risultati_elaborazione": [r.__dict__ for r in risultati],
            "timestamp": datetime.now().isoformat()
        }
        try:
            with open(self.file_path, 'w', encoding='utf-8') as f:
                json.dump(stato, f, indent=4, ensure_ascii=False)
        except Exception as e:
            logger.error(f"Impossibile salvare lo stato: {e}")

    def rimuovi_stato(self) -> None:
        """Elimina il file di stato alla fine di un processo completato."""
        if os.path.exists(self.file_path):
            os.remove(self.file_path)

class PDLOrchestrator:
    """Coordina l'intero processo di automazione."""

    def __init__(self, dry_run: bool = False, secure_pwd: bool = False) -> None:
        self.dry_run = dry_run
        self.secure_pwd = secure_pwd
        self.config_path = os.path.join(Config.SCRIPT_DIR, Config.EXCEL_FILE_CONFIG_NAME)
        self.excel = ExcelProcessor(self.config_path)
        self.driver_manager = WebDriverManager(headless=not dry_run, start_maximized=True)
        self.state = StateManager(os.path.join(Config.SCRIPT_DIR, Config.FILE_STATO_PROCESSO))

    def run(self) -> None:
        """Esegue il workflow principale."""
        logger.info(f"Avvio automazione (Dry Run: {self.dry_run})")
        
        try:
            # 1. Preparazione
            url = self.excel.get_website_url()
            user, pwd = self.excel.get_credentials(self.secure_pwd)
            
            if not self.dry_run:
                self.excel.run_pdl_macros()
            
            pdl_list = self.excel.get_pdl_list_from_excel()
            if not pdl_list:
                logger.info("Nessun PDL da processare oggi.")
                return

            # 2. Ripresa stato
            ultimo_idx, risultati_precedenti = self.state.carica_stato()
            # Uniamo i dati freschi da excel con eventuali stati già salvati
            # (In questa versione semplificata ripartiamo dalla lista excel completa)
            
            # 3. Ciclo di elaborazione con gestione errori sessione
            idx_corrente = ultimo_idx + 1
            tentativi_falliti = 0
            
            while idx_corrente < len(pdl_list):
                try:
                    logger.info(f"Apertura sessione browser per processare {len(pdl_list)-idx_corrente} PDL.")
                    driver = self.driver_manager.get_driver()
                    automator = SafeWorkAutomator(driver, dry_run=self.dry_run)
                    
                    automator.login(url, user, pwd)
                    automator.navigate_to_booking()
                    
                    for i in range(idx_corrente, len(pdl_list)):
                        pdl = pdl_list[i]
                        pdl.stato_script = automator.process_pdl(pdl)
                        
                        idx_corrente = i
                        self.state.salva_stato(i, pdl_list)
                        time.sleep(Config.PAUSA_TRA_PDL)
                        
                    break # Tutto completato
                    
                except (TimeoutAlertDetected, AutomationException) as e:
                    logger.error(f"Errore di sessione: {e}. Riavvio driver...")
                    tentativi_falliti += 1
                    self.driver_manager.quit_driver()
                    if tentativi_falliti >= Config.MAX_SETUP_ATTEMPTS:
                        logger.critical("Raggiunto numero massimo di tentativi di riavvio sessione.")
                        break
                    time.sleep(Config.PAUSA_TRA_TENTATIVI_SETUP_FALLITI)
            
            if idx_corrente >= len(pdl_list) - 1:
                logger.info("Processo completato con successo.")
                self.state.rimuovi_stato()

        except Exception as e:
            logger.critical(f"Errore fatale imprevisto: {e}")
            traceback.print_exc()
        finally:
            self.driver_manager.quit_driver()

def main() -> None:
    """Entry point CLI."""
    parser = argparse.ArgumentParser(description="Automazione SafeWork ISAB.")
    parser.add_argument("--dry-run", action="store_true", help="Simula le azioni senza scrivere sul sito.")
    parser.add_argument("--secure", action="store_true", help="Richiede la password interattivamente.")
    args = parser.parse_args()
    
    orchestrator = PDLOrchestrator(dry_run=args.dry_run, secure_pwd=args.secure)
    orchestrator.run()

if __name__ == "__main__":
    main()
