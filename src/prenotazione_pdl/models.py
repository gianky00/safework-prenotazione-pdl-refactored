"""Modelli dati ed eccezioni custom per il processo di prenotazione PDL."""

from dataclasses import dataclass
from typing import Optional

class AutomationException(Exception):
    """Eccezione base per gli errori di automazione."""
    pass

class CriticalConfigError(AutomationException):
    """Errore critico nella configurazione o nei parametri Excel."""
    pass

class ProcessInterruptedException(AutomationException):
    """Il processo è stato interrotto dall'utente o da un segnale esterno."""
    pass

class TimeoutAlertDetected(AutomationException):
    """È stato rilevato un alert di timeout dal sito web."""
    pass

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
