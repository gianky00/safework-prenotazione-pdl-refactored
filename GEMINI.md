# ♾️ SAFEWORK PDL - AI GUIDELINES (V1.0)

## 🚨 REGOLE FERREE
1. **LINGUA**: Rispondi sempre in ITALIANO.
2. **QUALITÀ**: Ogni modifica deve superare `ruff check` e `mypy`.
3. **DOCUMENTAZIONE**: Ogni nuova funzione deve avere docstring conformi a `interrogate` (fail-under: 90%).
4. **SURGICAL EDITS**: Modifiche mirate, ampie ancore di contesto.

## 🏗️ ARCHITETTURA
- Logica Selenium in `src/prenotazione_pdl/automation/`.
- Logica Excel in `src/prenotazione_pdl/excel/`.
- Modelli in `src/prenotazione_pdl/models.py`.
- Configurazione in `src/prenotazione_pdl/config.py`.
