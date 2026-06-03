# SafeWork Prenotazione PDL (Refactored)

Automazione professionale per la prenotazione dei Piani di Lavoro (PDL) sul portale SafeWork ISAB.

## 🚀 Panoramica
Questa è la versione rifattorizzata e modulare dello script originale, progettata per essere robusta, tipizzata e facilmente manutenibile.

### Caratteristiche principali
- **Architettura Modulare**: Separazione netta tra logica Selenium, gestione Excel e orchestrazione.
- **Qualità Enterprise**: Configurazione completa per `ruff`, `mypy`, `interrogate`, `xenon` e `radon`.
- **Robustezza**: Gestione avanzata dei timeout del sito e ripristino automatico delle sessioni browser.
- **Persistenza**: Salvataggio dello stato di avanzamento per riprendere l'elaborazione in caso di interruzioni.

## 🛠️ Installazione
Il progetto utilizza **Poetry** per la gestione delle dipendenze.

```bash
# Installa le dipendenze
poetry install

# Oppure via pip
pip install -r requirements.txt
```

## 📖 Utilizzo
L'entry point principale è `src/main.py`.

```bash
# Esecuzione standard
python src/main.py

# Esecuzione in modalità simulazione (senza salvare sul sito)
python src/main.py --dry-run

# Richiesta password interattiva (più sicuro)
python src/main.py --secure
```

## 💎 Qualità del Codice
Il progetto segue standard rigorosi. I controlli vengono eseguiti automaticamente ad ogni commit tramite **pre-commit**:

```bash
# Linter e Formatter
ruff check src
ruff format src

# Type checking (Ottimizzato per Windows/Linux)
python -m mypy src

# Documentazione
interrogate src
```

## 💎 Interfaccia Utente & Grafica
L'applicazione vanta un'interfaccia CLI moderna e professionale basata sulla libreria **Rich**, allineata agli standard estetici degli strumenti RPA SafeWork:
- **Logo ASCII Dinamico**: Branding coerente con sfumature ciano/blu.
- **Progress Tracking**: Barre di avanzamento vettoriali con calcolo del tempo rimanente.
- **Reporting**: Tabelle stilizzate e pannelli informativi per un feedback immediato.
- **Logging Avanzato**: Integrazione tra `Loguru` e `RichHandler` per log colorati e leggibili direttamente in console.

## 🏗️ Struttura Progetto
- `src/`: Codice sorgente Python.
  - `automation/`: Driver Selenium e azioni sul portale.
  - `excel/`: Lettura parametri ed esecuzione macro.
  - `utils/`: Gestione email e stampa professionale.
  - `main.py`: Orchestratore principale.
- `data/`: Archiviazione centralizzata.
  - `logs/`: Storico esecuzioni gerarchico per data.
  - `reports/`: PDF professionali generati.
  - `database/`: Persistenza dati SQLite.
  - `state/`: File di ripristino sessione.
  - `parametri prenotazione pdl.xlsx`: Pannello di controllo Excel.
- `tests/`: Suite di test (pytest).
