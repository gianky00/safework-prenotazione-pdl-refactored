"""Script di test rapido per verificare esclusivamente il layout PDF professionale."""

import os
import sys
from datetime import datetime

from loguru import logger

# Aggiungo la cartella src al path per poter importare i moduli
sys.path.append(os.path.join(os.path.dirname(__file__), "src"))

from src.models import PDLData
from src.utils.printer_manager import PrinterManager


def test_layout():
    """Esegue il test di stampa del layout PDF."""
    logger.info("Avvio TEST LAYOUT PDF (Senza stampa fisica)...")

    # 1. Creazione dati di test (fittizi e variati per testare il layout)
    pdl_test = [
        PDLData(riga_excel_debug=10, pdl="586559/C", area="AREA 3 - RAFFINERIA", impianto="700 - TOPPING", tempo_rimanente="09:51:32 (1 ore 54 minuti )"),
        PDLData(riga_excel_debug=11, pdl="587028/C", area="AREA 3 - RAFFINERIA", impianto="1200 - ISOMERIZZAZIONE", tempo_rimanente="09:54:32 (1 ore 57 minuti )"),
        PDLData(riga_excel_debug=12, pdl="586751/C", area="AREA 3 - RAFFINERIA", impianto="G.O.D.", tempo_rimanente="09:57:32 (2 ore )"),
        PDLData(riga_excel_debug=13, pdl="587026/C", area="AREA 3", impianto="2200 BD", tempo_rimanente="10:00:32 (2 ore 3 minuti )"),
        PDLData(riga_excel_debug=14, pdl="586246/C", area="AREA 4 - ESTERNO", impianto="PONTILI", tempo_rimanente="10:15:00 (In corso)")
    ]

    # Impostiamo degli stati per vedere i diversi stili
    pdl_test[0].stato_script = "SUCCESSO"
    pdl_test[1].stato_script = "GIÀ PRENOTATO"
    pdl_test[2].stato_script = "PENDING - VERIFICA"
    pdl_test[3].stato_script = "ERRORE CONNESSIONE"
    pdl_test[4].stato_script = "IN ELABORAZIONE..."

    # 2. Inizializzazione PrinterManager
    pm = PrinterManager()
    project_root = os.path.dirname(os.path.abspath(__file__))

    # 3. Generazione manuale del PDF (Bypassing la stampa fisica)
    archive_dir = os.path.join(project_root, "report_pdf", "test_visuale")
    os.makedirs(archive_dir, exist_ok=True)

    file_path = os.path.join(archive_dir, "Test_Layout_Coemi_Premium.pdf")

    try:
        logger.info(f"Generazione PDF in corso: {file_path}")
        pm._generate_pdf(pdl_test, file_path, project_root, datetime.now())

        logger.success("PDF generato con successo!")
        print("\n--- TEST LAYOUT COMPLETATO ---")
        print(f"File salvato in: {file_path}")
        print("Apertura automatica del file per verifica layout...")

        # Apertura automatica del PDF su Windows
        os.startfile(file_path)

    except Exception as e:
        logger.error(f"Errore durante il test di layout: {e}")

if __name__ == "__main__":
    test_layout()
