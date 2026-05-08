# Piano di Rifattorizzazione: SafeWork Prenotazione PDL

Il progetto verrà rifattorizzato per migliorare la manutenibilità, la leggibilità e l'estensibilità, seguendo gli standard di qualità del progetto `ISAB_TimeSheet`.

## 1. Obiettivi
- **Modularizzazione**: Suddividere il monolite `SafeWorkPrenotaPDL.py` in moduli logici.
- **Qualità del Codice**: Configurare e applicare `ruff`, `mypy`, `interrogate`, `xenon`, `radon`.
- **Tipizzazione**: Aggiungere type hints completi per il controllo statico.
- **GitHub**: Creare un nuovo repository e versionare il codice rifattorizzato.

## 2. Nuova Struttura Progetto (src Layout)
```text
prenotazione_pdl/
├── .github/workflow/       # CI/CD (opzionale)
├── src/
│   └── prenotazione_pdl/
│       ├── __init__.py
│       ├── main.py         # Entry point
│       ├── config.py       # Configurazioni e costanti
│       ├── models.py       # Dataclass e modelli dati
│       ├── automation/     # Logica Selenium
│       │   ├── __init__.py
│       │   ├── driver.py   # Gestione WebDriver
│       │   └── actions.py  # Azioni sul sito SafeWork
│       └── excel/          # Gestione Excel e Macro
│           ├── __init__.py
│           └── processor.py
├── tests/                  # Test unitari e integrazione
├── pyproject.toml          # Configurazione strumenti
├── requirements.txt        # Dipendenze
└── SafeWorkPrenotaPDL.py   # Vecchio codice (mantenuto)
```

## 3. Fasi di Esecuzione

### Fase 1: Setup Ambiente
1. [ ] Creazione repository GitHub.
2. [ ] Inizializzazione Git locale.
3. [ ] Creazione `pyproject.toml` basato su `ISAB_TimeSheet`.
4. [ ] Creazione struttura directory `src/`.

### Fase 2: Estrazione Moduli
1. [ ] **Config**: Spostare `Config` in `src/prenotazione_pdl/config.py`.
2. [ ] **Models**: Spostare `PDLData` e eccezioni in `src/prenotazione_pdl/models.py`.
3. [ ] **Excel**: Rifattorizzare `ExcelProcessor` in `src/prenotazione_pdl/excel/processor.py`.
4. [ ] **WebDriver**: Rifattorizzare `WebDriverManager` in `src/prenotazione_pdl/automation/driver.py`.
5. [ ] **Automation**: Rifattorizzare `SafeWorkAutomator` in `src/prenotazione_pdl/automation/actions.py`.
6. [ ] **Orchestrator**: Rifattorizzare `PDLOrchestrator` e `main` in `src/prenotazione_pdl/main.py`.

### Fase 3: Qualità e Validazione
1. [ ] Esecuzione `ruff check` e `ruff format`.
2. [ ] Controllo tipi con `mypy`.
3. [ ] Verifica documentazione con `interrogate`.
4. [ ] Analisi complessità con `xenon` e `radon`.
5. [ ] Esecuzione test (se presenti/creati).

### Fase 4: Consegna
1. [ ] Push iniziale sul nuovo repository GitHub.
