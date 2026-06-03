# CLAUDE.md

## Project Overview
**SafeWork Prenotazione PDL** è un'automazione Selenium per il portale SafeWork ISAB.

## Development Commands
```bash
# Setup
poetry install
pip install -e .

# Qualità (Eseguiti automaticamente via pre-commit)
ruff check src
python -m mypy src
ruff format src
interrogate src

# Esecuzione
python src/main.py
```

## Architecture
- **src Layout**: Moduli logici organizzati per responsabilità.
- **Data & Logs**: Archiviazione centralizzata in `data/` (logs, reports, database, config).
- **Orchestrator Pattern**: `PDLOrchestrator` gestisce il flusso principale.
- **Driver Management**: `WebDriverManager` incapsula la creazione di Chrome.
- **Excel Interop**: `ExcelProcessor` gestisce l'integrazione con i parametri Excel.
