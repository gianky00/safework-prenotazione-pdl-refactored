# Piano di Rifattorizzazione: SafeWork Prenotazione PDL

Il progetto è stato rifattorizzato per migliorare la manutenibilità, la leggibilità e l'estensibilità, seguendo gli standard di qualità del progetto `ISAB_TimeSheet`.

## 1. Obiettivi
- **Modularizzazione**: Suddividere il monolite `SafeWorkPrenotaPDL.py` in moduli logici direttamente sotto `src/` (layout a monte).
- **Qualità del Codice**: Configurare e applicare `ruff` e `mypy`.
- **Tipizzazione**: Aggiungere type hints completi per il controllo statico.
- **Automazione Invisibile**: Abilitazione della modalità headless di default per l'esecuzione in background.
- **Riorganizzazione Root**: Pulizia dei file obsoleti e rinomina dell'eseguibile in `Avvia.bat`.

## 2. Struttura Progetto Attuale (src Layout a Monte)
```text
prenotazione_pdl/
├── .git/
├── .github/
├── archive/
│   └── SafeWorkPrenotaPDL.py   # Vecchio codice monolitico archiviato
├── src/                        # Codice sorgente a monte
│   ├── __init__.py
│   ├── main.py                 # Entry point (Orchestrator)
│   ├── config.py               # Configurazioni e costanti (risoluzione a 2 livelli)
│   ├── models.py               # Dataclass e modelli dati
│   ├── automation/             # Logica Selenium
│   │   ├── __init__.py
│   │   ├── driver.py           # Gestione WebDriver
│   │   └── actions.py          # Azioni sul portale (con import WebElement)
│   └── excel/                  # Gestione Excel e Macro
│       ├── __init__.py
│       └── processor.py
├── tests/                      # Test unitari e integrazione
├── pyproject.toml              # Configurazione strumenti (Poetry, Ruff, Mypy)
├── requirements.txt            # Dipendenze
├── parametri prenotazione pdl.xlsx # File parametri Excel nella root
├── Avvia.bat                   # Eseguibile batch principale (headless di default)
└── prenotazione_pdl.log        # File di log attivo
```

## 3. Fasi di Esecuzione

### Fase 1: Setup Ambiente
1. [x] Creazione repository GitHub.
2. [x] Inizializzazione Git locale.
3. [x] Creazione `pyproject.toml` basato su `ISAB_TimeSheet`.
4. [x] Creazione struttura directory `src/`.

### Fase 2: Estrazione Moduli (Spostati a Monte)
1. [x] **Config**: Spostare `Config` in `src/config.py` (con supporto salita a 2 livelli).
2. [x] **Models**: Spostare `PDLData` e eccezioni in `src/models.py`.
3. [x] **Excel**: Rifattorizzare `ExcelProcessor` in `src/excel/processor.py`.
4. [x] **WebDriver**: Rifattorizzare `WebDriverManager` in `src/automation/driver.py`.
5. [x] **Automation**: Rifattorizzare `SafeWorkAutomator` in `src/automation/actions.py` (con import `WebElement`).
6. [x] **Orchestrator**: Rifattorizzare `PDLOrchestrator` e `main` in `src/main.py`.

### Fase 3: Qualità e Validazione
1. [x] Esecuzione `ruff check` e `ruff check --fix` (import ordinati).
2. [x] Controllo tipi con `mypy` e risoluzione dei tipi mancanti.
3. [x] Risoluzione del bug di importazione di `WebElement` in `actions.py`.
4. [x] Integrazione ed esecuzione in modalità **headless** di default con opzione CLI `--no-headless`.
5. [x] Rinomina dell'eseguibile batch in `Avvia.bat` e verifica di bootstrapping.

### Fase 4: Consegna e Pulizia
1. [x] Creazione directory `archive/` e archiviazione di `SafeWorkPrenotaPDL.py`.
2. [x] Rimozione dei file obsoleti (`CRITICAL_EXECUTION_ERROR.txt`, `test.db` e log residui).
3. [x] Push iniziale sul nuovo repository GitHub.
