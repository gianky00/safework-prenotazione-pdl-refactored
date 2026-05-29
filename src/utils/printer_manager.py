"""Modulo per la generazione e stampa di report PDF professionali su Windows."""

import os
from datetime import datetime

import win32api
import win32print
from loguru import logger
from reportlab.lib import colors  # type: ignore
from reportlab.lib.pagesizes import A4  # type: ignore
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet  # type: ignore
from reportlab.lib.units import cm  # type: ignore
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle  # type: ignore

from .email_manager import PDLData


class PrinterManager:
    """Gestisce la creazione di PDF professionali e l'invio alla stampante di sistema."""

    def __init__(self, printer_name: str | None = None) -> None:
        """Inizializza il manager della stampante."""
        self.printer_name: str | None = None
        try:
            self.printer_name = printer_name or win32print.GetDefaultPrinter()
        except Exception:
            logger.warning("Nessuna stampante predefinita trovata.")

    def print_pdl_report(self, pdl_list: list[PDLData]) -> bool:
        """Genera un PDF professionale, lo salva in archivio e lo invia alla stampante."""
        if not pdl_list:
            logger.warning("Nessun dato da stampare.")
            return False

        try:
            # Definizione cartella e nome file professionale
            base_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
            archive_dir = os.path.join(base_dir, "report_pdf")
            os.makedirs(archive_dir, exist_ok=True)

            timestamp_file = datetime.now().strftime("%Y-%m-%d_%H-%M")
            file_name = f"Report_PDL_{timestamp_file}.pdf"
            file_path = os.path.join(archive_dir, file_name)

            # Generazione contenuto PDF
            self._generate_pdf(pdl_list, file_path)
            logger.info(f"Report PDF archiviato: {file_path}")

            if not self.printer_name:
                logger.error("Impossibile stampare: stampante non configurata.")
                return True # Ritorna comunque True perché il salvataggio è riuscito

            logger.info(f"Invio PDF alla stampante: {self.printer_name}")

            # Comando di stampa nativo Windows
            win32api.ShellExecute(
                0,
                "print",
                file_path,
                f'/d:"{self.printer_name}"',
                ".",
                0
            )
            return True

        except Exception as e:
            logger.error(f"Errore durante l'archiviazione/stampa PDF: {e}")
            return False

    def _generate_pdf(self, pdl_list: list[PDLData], file_path: str) -> None:
        """Costruisce il layout del PDF professionale."""
        doc = SimpleDocTemplate(
            file_path,
            pagesize=A4,
            rightMargin=1.5*cm,
            leftMargin=1.5*cm,
            topMargin=1.5*cm,
            bottomMargin=1.5*cm
        )

        elements = []
        styles = getSampleStyleSheet()

        # --- TITOLO ---
        title_style = ParagraphStyle(
            'TitleStyle',
            parent=styles['Heading1'],
            fontSize=18,
            alignment=1,  # Center
            spaceAfter=12,
            textColor=colors.black
        )
        elements.append(Paragraph("REPORT PRENOTAZIONE PDL", title_style))

        # --- SOTTOTITOLO (DATA) ---
        sub_style = ParagraphStyle(
            'SubStyle',
            parent=styles['Normal'],
            fontSize=10,
            alignment=1,
            spaceAfter=20,
            textColor=colors.grey
        )
        now_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        elements.append(Paragraph(f"Documento generato il {now_str}", sub_style))

        # --- TABELLA DATI ---
        # Intestazioni
        data = [['PdL', 'Area', 'Impianto', 'Orario prenotazione', 'Esito']]

        # Righe PdL
        for pdl in pdl_list:
            # Pulisce il tempo togliendo eventuali scritte tra parentesi (es. "09:45:23 (2 ore...)")
            tempo_pulito = str(pdl.tempo_rimanente or "-").split('(')[0].strip()

            data.append([
                str(pdl.pdl),
                str(pdl.area),
                str(pdl.impianto),
                tempo_pulito,
                str(pdl.stato_script)
            ])

        # Definizione larghezze colonne (A4 è circa 21cm, tolti i margini rimangono 18cm)
        col_widths = [3.5*cm, 3.5*cm, 3.5*cm, 4*cm, 3.5*cm]
        table = Table(data, colWidths=col_widths, repeatRows=1)

        # Stile Tabella (Ottimizzato B/N)
        table_style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),  # Intestazione grigio chiaro
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 10),

            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),      # Bordi neri sottili
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('TOPPADDING', (0, 1), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
        ])

        # Evidenziazione righe alternate (opzionale per B/N)
        for i in range(1, len(data)):
            if i % 2 == 0:
                table_style.add('BACKGROUND', (0, i), (-1, i), colors.whitesmoke)

        table.setStyle(table_style)
        elements.append(table)

        # --- FOOTER ---
        elements.append(Spacer(1, 2*cm))
        footer_style = ParagraphStyle(
            'FooterStyle',
            parent=styles['Normal'],
            fontSize=8,
            alignment=2,  # Right
            textColor=colors.grey
        )
        elements.append(Paragraph("Sistema Automazione SafeWork ISAB Sud", footer_style))

        # Salvataggio
        doc.build(elements)
        logger.info(f"PDF generato con successo: {file_path}")
