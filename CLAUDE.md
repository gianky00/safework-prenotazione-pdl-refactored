# CLAUDE.md

## Project Overview
**SafeWork Prenotazione PDL** è un'automazione Selenium per il portale SafeWork ISAB.

## Development Commands
```bash
# Setup
poetry install
pip install -e .

# Qualità
ruff check src
mypy src
ruff format src
interrogate src

# Esecuzione
python src/prenotazione_pdl/main.py
```

## Architecture
- **src Layout**: Segue la struttura moderna dei pacchetti Python.
- **Orchestrator Pattern**: `PDLOrchestrator` gestisce il flusso principale.
- **Driver Management**: `WebDriverManager` incapsula la creazione di Chrome.
- **Excel Interop**: `ExcelProcessor` gestisce l'integrazione con i parametri Excel.
