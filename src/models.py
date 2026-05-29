"""Modelli dati ed eccezioni custom per il processo di prenotazione PDL."""

from dataclasses import dataclass


class AutomationError(Exception):
    """Eccezione base per gli errori di automazione."""


class CriticalConfigError(AutomationError):
    """Errore critico nella configurazione o nei parametri Excel."""


class ProcessInterruptedError(AutomationError):
    """Il processo è stato interrotto dall'utente o da un segnale esterno."""


class TimeoutAlertError(AutomationError):
    """È stato rilevato un alert di timeout dal sito web."""


@dataclass
class PDLData:
    """Modello per i dati di un singolo PDL estratti da Excel."""

    riga_excel_debug: int
    pdl: str = ""
    area: str = "Area Non Specificata"
    descrizione: str = ""
    stato_pdl_excel: str = ""
    stato_attivita_excel: str = ""
    data_controllo_excel: str = ""
    personale_excel: str = ""
    impianto: str = ""
    stato_script: str = "Non Processato"
    tempo_rimanente: str = ""
