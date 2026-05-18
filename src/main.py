"""Entry point principale per l'automazione Prenotazione PDL."""

import argparse
import contextlib
import json
import msvcrt
import os
import sys
import time
import traceback
import warnings
from datetime import UTC, datetime
from typing import Any

from loguru import logger
from rich import box
from rich.align import Align
from rich.console import Console
from rich.logging import RichHandler
from rich.panel import Panel
from rich.progress import (
    BarColumn,
    Progress,
    SpinnerColumn,
    TaskProgressColumn,
    TextColumn,
    TimeRemainingColumn,
)
from rich.rule import Rule
from rich.table import Table
from rich.text import Text
from rich.theme import Theme

from .automation.actions import SafeWorkAutomator
from .automation.driver import WebDriverManager
from .config import Config
from .excel.processor import ExcelProcessor
from .models import (
    AutomationError,
    PDLData,
    TimeoutAlertError,
)

# SOPPRESSIONE WARNING E LOG DI SISTEMA
os.environ['TF_CPP_MIN_LOG_LEVEL'] = '3'
os.environ['WDM_LOG_LEVEL'] = '0'
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Inizializzazione Console Rich con Tema Personalizzato
custom_theme = Theme({
    "info": "#00d2ff",
    "warning": "#ff8c00",
    "error": "bold #ff3333",
    "success": "bold #00ff7f",
    "brand": "bold #0088ff",
    "highlight": "bold #ffd700",
    "muted": "dim #888888"
})
console = Console(theme=custom_theme)


def stampa_logo() -> None:
    """Stampa il logo ASCII dell'applicazione."""
    logo_text = r"""
[#00e5ff]   ______ ____  ______ __  __ ____ [/#00e5ff]
[#00ccff]  / ____// __ \/ ____//  |/  //  _/ [/#00ccff]
[#00b3ff] / /    / / / / __/  / /|_/ / / /   [/#00b3ff]
[#0099ff]/ /___ / /_/ / /___ / /  / /_/ /      ___  ___   _   [/#0099ff]
[#0080ff]\____/ \____/_____//_/  /_//___/     / __|| _ \ | |  [/#0080ff]
[#0066ff]                                     \__ \|   / | |__[/#0066ff]
[#004dff]                                     |___/|_|_\ |____|[/#004dff]

[highlight]🤖 SAFEWORK PRENOTAZIONE PDL v2.0[/highlight]
[muted]👨‍💻 Developer: Giancarlo Allegretti  |  📅 Revision: 08/05/2026[/muted]"""

    panel_logo = Panel(
        Text.from_markup(logo_text, justify="center"),
        border_style="#00d2ff",
        padding=(1, 4),
        style="on #000B1A",
        title="[bold white]PRENOTAZIONE AUTOMATICA PDL[/bold white]",
        box=box.DOUBLE_EDGE,
        expand=False
    )
    console.print(Align.center(panel_logo))
    console.print("\n")


def timed_input(prompt: str, timeout: int = 3600) -> str | None:
    """Attende un input dall'utente per un tempo massimo (Windows)."""
    sys.stdout.write(prompt)
    sys.stdout.flush()
    start_time = time.time()
    input_str = ""
    while (time.time() - start_time) <= timeout:
        if msvcrt.kbhit():
            res = _gestisci_carattere_input(input_str)
            if res is True:
                return input_str.strip()
            input_str = str(res)
        time.sleep(0.05)

    sys.stdout.write("\n[TEMPO SCADUTO] Proseguo automaticamente...\n")
    return None


def _gestisci_carattere_input(current_str: str) -> str | bool:
    """Gestisce la lettura di un singolo carattere da console (Windows)."""
    char = msvcrt.getwche()
    if char in ("\r", "\n"):
        sys.stdout.write("\n")
        return True
    if char == "\b":
        if len(current_str) > 0:
            sys.stdout.write(" \b")
            return current_str[:-1]
        sys.stdout.write(" ")
        return current_str
    return current_str + char


def _get_esito_styled(esito: str) -> str:
    """Restituisce l'esito formattato con i colori Rich."""
    esito_upper = esito.upper()
    success_keywords = ["OK", "COMPLETATO", "ESEGUITA"]
    error_keywords = ["ERR", "FALLITO"]

    if any(k in esito_upper for k in success_keywords):
        return f"[success]{esito}[/success]"
    if any(k in esito_upper for k in error_keywords):
        return f"[error]{esito}[/error]"
    return f"[warning]{esito}[/warning]"


def stampa_report_finale(pdl_list: list[PDLData]) -> None:
    """Genera una tabella riassuntiva dei risultati in stile Rich."""
    if not pdl_list:
        return

    table = Table(
        title=Text("📊 RIEPILOGO ELABORAZIONE PDL", justify="center", style="#ffffff bold"),
        show_header=True,
        header_style="bold #000000 on #00d2ff",
        border_style="#0088ff",
        box=box.ROUNDED,
        expand=True,
        row_styles=["none"],
        caption=f"Totale processati: [bold highlight]{len(pdl_list)}[/bold highlight]",
        caption_style="italic muted",
    )

    table.add_column("PdL", style="success", no_wrap=True)
    table.add_column("AREA", style="#b3e6ff")
    table.add_column("IMPIANTO", style="#e6f7ff")
    table.add_column("ESITO AUTOMAZIONE", justify="center")

    for pdl in pdl_list:
        esito = pdl.stato_script or "NON PROCESSATO"
        table.add_row(str(pdl.pdl), str(pdl.area), str(pdl.impianto), _get_esito_styled(esito))

    console.print("\n", table, "\n")
    _stampa_pannello_recap(pdl_list)


def _stampa_pannello_recap(pdl_list: list[PDLData]) -> None:
    """Calcola le statistiche e stampa il pannello di recap finale."""
    success_keys = ["OK", "COMPLETATO", "ESEGUITA"]
    successi = sum(
        1
        for p in pdl_list
        if p.stato_script and any(k in p.stato_script.upper() for k in success_keys)
    )
    anomalie = len(pdl_list) - successi

    summary_text = f"""[#00d2ff]🔹 PdL Totali in lista:[/#00d2ff] [highlight]{len(pdl_list)}[/highlight]
[#00d2ff]🔹 Prenotazioni completate:[/#00d2ff] [success]{successi}[/success]
[#00d2ff]🔹 Anomalie rilevate:[/#00d2ff] [error]{anomalie}[/error]"""

    panel = Panel(
        summary_text,
        title="[highlight]📊 RECAP SESSIONE[/highlight]",
        padding=(1, 5),
        border_style="#ffd700",
        box=box.HEAVY,
        expand=False,
        style="on #1a1a00",
    )
    console.print(Align.center(panel), "\n")


class StateManager:
    """Gestisce la persistenza dello stato dell'elaborazione in un file JSON."""

    def __init__(self, file_path: str) -> None:
        self.file_path = file_path

    def carica_stato(self) -> tuple[int, list[PDLData]]:
        """Carica l'ultimo indice processato e i risultati parziali."""
        if not self._is_stato_valido():
            return -1, []

        try:
            with open(self.file_path, encoding="utf-8") as f:
                stato = json.load(f)
            idx = stato.get("ultimo_indice_pdl_processato", -1)
            dati = [PDLData(**d) for d in stato.get("risultati_elaborazione", [])]
            return idx, dati
        except Exception as e:
            logger.warning(f"Errore caricamento stato ({e}). Ripristino...")
            self.rimuovi_stato()
            return -1, []

    def _is_stato_valido(self) -> bool:
        """Verifica se il file di stato esiste ed è popolato."""
        return os.path.exists(self.file_path) and os.path.getsize(self.file_path) > 0

    def salva_stato(self, ultimo_indice: int, risultati: list[PDLData]) -> None:
        """Salva lo stato corrente su disco."""
        stato = {
            "timestamp": datetime.now(UTC).isoformat(),
            "ultimo_indice_pdl_processato": ultimo_indice,
            "risultati_elaborazione": [r.__dict__ for r in risultati]
        }
        try:
            with open(self.file_path, 'w', encoding='utf-8') as f:
                json.dump(stato, f, indent=4)
        except Exception as e:
            logger.error(f"Impossibile salvare lo stato: {e}")

    def rimuovi_stato(self) -> None:
        """Elimina il file di stato."""
        if os.path.exists(self.file_path):
            with contextlib.suppress(Exception):
                os.remove(self.file_path)


class PDLOrchestrator:
    """Orchestra l'intero workflow di prenotazione PDL."""

    def __init__(self, dry_run: bool = False, secure_pwd: bool = False, headless: bool = True) -> None:
        self.dry_run = dry_run
        self.secure_pwd = secure_pwd
        self.headless = headless

        # Configurazione Logger immediata
        logger.remove()
        logger.add(
            RichHandler(console=console, rich_tracebacks=True, markup=True, show_time=True, show_path=False),
            format="{message}",
            level="SUCCESS"
        )
        logger.add(
            os.path.join(Config.SCRIPT_DIR, "prenotazione_pdl.log"),
            format="{time:YYYY-MM-DD HH:mm:ss} | {level: <8} | {message}",
            mode="w",
            level="DEBUG"
        )

        self.config_path = os.path.join(Config.SCRIPT_DIR, Config.EXCEL_FILE_CONFIG_NAME)
        self.excel = ExcelProcessor(self.config_path)
        self.driver_manager = WebDriverManager(headless=self.headless, start_maximized=True)
        self.state = StateManager(os.path.join(Config.SCRIPT_DIR, Config.FILE_STATO_PROCESSO))

    def _inizializza_dati_preparazione(self, progress: Progress, task_id: Any) -> tuple[str, str, str, list[PDLData]]:
        """Esegue la fase 1: recupero URL, credenziali e lista PDL."""
        progress.update(task_id, description="[info]📖 Recupero URL e credenziali...[/info]")
        url = self.excel.get_website_url()
        user, pwd = self.excel.get_credentials(self.secure_pwd)

        if not self.dry_run:
            progress.update(task_id, description="[info]🚀 Esecuzione macro Excel...[/info]")
            self.excel.run_pdl_macros()

        progress.update(task_id, description="[info]📋 Estrazione lista PdL...[/info]")
        pdl_list = self.excel.get_pdl_list_from_excel()
        return url, user, pwd, pdl_list

    def _elabora_pdl_loop(
        self,
        progress: Progress,
        task_id: Any,
        pdl_list: list[PDLData],
        idx_corrente: int,
        url: str,
        user: str,
        pwd: str,
    ) -> int:
        """Gestisce il loop principale di elaborazione dei PDL con gestione dei riavvii driver."""
        tentativi_falliti = 0
        while idx_corrente < len(pdl_list):
            try:
                progress.update(task_id, description="[info]🌐 Accesso al portale in corso...[/info]")
                driver = self.driver_manager.get_driver()
                automator = SafeWorkAutomator(driver, dry_run=self.dry_run)

                automator.login(url, user, pwd)
                automator.navigate_to_booking()

                for i in range(idx_corrente, len(pdl_list)):
                    pdl = pdl_list[i]
                    progress.update(task_id, description=f"[info]🚀 Processo PdL: [highlight]{pdl.pdl}[/highlight][/info]")
                    pdl.stato_script = automator.process_pdl(pdl)
                    idx_corrente = i + 1
                    self.state.salva_stato(i, pdl_list)
                    progress.advance(task_id)
                    time.sleep(Config.PAUSA_TRA_PDL)
                break

            except (TimeoutAlertError, AutomationError) as e:
                logger.error(f"Errore di sessione: {e}. Riavvio driver...")
                tentativi_falliti += 1
                self.driver_manager.quit_driver()
                if tentativi_falliti >= Config.MAX_SETUP_ATTEMPTS:
                    console.print("[error]❌ Raggiunto numero massimo di tentativi di riavvio.[/error]")
                    break
                time.sleep(Config.PAUSA_TRA_TENTATIVI_SETUP_FALLITI)
        return idx_corrente

    def run(self) -> None:
        """Esegue il workflow principale con un'unica barra di avanzamento consolidata."""
        stampa_logo()
        logger.info(f"Avvio automazione (Dry Run: {self.dry_run})")

        with Progress(
            SpinnerColumn(spinner_name="point", style="highlight"),
            TextColumn("[progress.description]{task.description}"),
            BarColumn(bar_width=None, style="#002b40", complete_style="#00d2ff", pulse_style="brand"),
            TaskProgressColumn(text_format="[progress.percentage]{task.percentage:>3.0f}%"),
            TimeRemainingColumn(),
            console=console,
            transient=False,
        ) as progress:
            try:
                # --- FASE 1: PREPARAZIONE ---
                console.print(Rule("[bold white]FASE 1: PREPARAZIONE[/bold white]", characters="▓▒░", style="#00d2ff"))
                prep_task = progress.add_task("[info]📖 Inizializzazione dati...[/info]", total=None)
                url, user, pwd, pdl_list = self._inizializza_dati_preparazione(progress, prep_task)
                progress.remove_task(prep_task)

                if not pdl_list:
                    console.print(Panel("[warning]⚠️ Nessun PDL da processare oggi.[/warning]", border_style="warning"))
                    return

                # --- FASE 2: ELABORAZIONE SUL PORTALE ---
                console.print("\n")
                console.print(
                    Rule("[bold white]FASE 2: ELABORAZIONE SUL PORTALE[/bold white]", characters="▓▒░", style="#00d2ff")
                )
                ultimo_idx, _ = self.state.carica_stato()
                idx_corrente = ultimo_idx + 1
                main_task = progress.add_task(
                    "[info]⚡ Inizializzazione sessione SafeWork...[/info]", total=len(pdl_list)
                )
                progress.update(main_task, completed=idx_corrente)

                idx_corrente = self._elabora_pdl_loop(progress, main_task, pdl_list, idx_corrente, url, user, pwd)

                # --- FASE 3: REPORTING ---
                if idx_corrente >= len(pdl_list):
                    progress.stop()
                    console.print("\n")
                    console.print(Rule("[bold white]FASE 3: REPORTING[/bold white]", characters="▓▒░", style="#00d2ff"))
                    stampa_report_finale(pdl_list)
                    self.state.rimuovi_stato()

                timed_input("\nOperazione conclusa. Premi INVIO per uscire (timeout 30m)... ", 1800)

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
    parser.add_argument("--no-headless", dest="headless", action="store_false", help="Mostra l'interfaccia grafica del browser.")
    parser.set_defaults(headless=True)
    args = parser.parse_args()

    orchestrator = PDLOrchestrator(dry_run=args.dry_run, secure_pwd=args.secure, headless=args.headless)
    orchestrator.run()


if __name__ == "__main__":
    main()
