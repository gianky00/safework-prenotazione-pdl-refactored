"""Modulo per la generazione e stampa di report PDF professionali su Windows."""

import os
from datetime import datetime
from typing import Any

import win32api
import win32print
from loguru import logger
from reportlab.lib import colors  # type: ignore
from reportlab.lib.pagesizes import A4  # type: ignore
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet  # type: ignore
from reportlab.lib.units import cm  # type: ignore
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle  # type: ignore

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
            # Definizione cartella e nome file professionale con gerarchia Anno/Mese
            project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

            now = datetime.now()
            year_str = now.strftime("%Y")
            month_str = now.strftime("%m")

            archive_dir = os.path.join(project_root, "report_pdf", year_str, month_str)
            os.makedirs(archive_dir, exist_ok=True)

            # Formato data richiesto: GG-MM-AAAA_HH-MM
            timestamp_file = now.strftime("%d-%m-%Y_%H-%M")
            file_name = f"Report_PDL_{timestamp_file}.pdf"
            file_path = os.path.join(archive_dir, file_name)

            # Generazione contenuto PDF
            self._generate_pdf(pdl_list, file_path, project_root, now)
            logger.info(f"Report PDF archiviato: {file_path}")

            if not self.printer_name:
                logger.error("Impossibile stampare: stampante non configurata.")
                return True

            logger.info(f"Tentativo di invio PDF alla stampante: {self.printer_name}")

            # Strategia di stampa robusta per Windows
            # 1. Tenta via Adobe Acrobat Reader (se presente) con flag silenti
            acrobat_path = r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
            if os.path.exists(acrobat_path):
                try:
                    # Flag /t: stampa su stampante specifica e chiude
                    logger.info("Utilizzo Adobe Acrobat per stampa silente...")
                    win32api.ShellExecute(
                        0,
                        "open",
                        acrobat_path,
                        f'/t "{file_path}" "{self.printer_name}"',
                        ".",
                        0
                    )
                    return True
                except Exception as e_acro:
                    logger.warning(f"Stampa via Acrobat fallita: {e_acro}")

            # 2. Fallback: Tentativo ShellExecute 'print' standard
            try:
                win32api.ShellExecute(0, "print", file_path, f'/d:"{self.printer_name}"', ".", 0)
                return True
            except Exception as e_shell:
                logger.warning(f"Metodo ShellExecute 'print' fallito ({e_shell}). Tento apertura documento.")
                win32api.ShellExecute(0, "open", file_path, "", ".", 1)
                return True

        except Exception as e:
            logger.error(f"Errore durante l'archiviazione/stampa PDF: {e}")
            return False

    def _generate_pdf(self, pdl_list: list[PDLData], file_path: str, project_root: str, now: datetime) -> None:
        """Costruisce il layout del PDF professionale in bianco e nero."""
        max_status_len = 30
        doc = SimpleDocTemplate(
            file_path,
            pagesize=A4,
            rightMargin=1.5 * cm,
            leftMargin=1.5 * cm,
            topMargin=1.2 * cm,
            bottomMargin=1.2 * cm
        )

        elements = []
        styles = getSampleStyleSheet()

        # --- 1. HEADER CON LOGO ---
        logo_path = os.path.join(project_root, "assets", "logo coemi.png")

        # Metadata del documento
        now_str = now.strftime("%d/%m/%Y %H:%M")

        meta_style = ParagraphStyle(
            'MetaStyle',
            fontSize=8,
            textColor=colors.grey,
            alignment=2  # Destra
        )

        disclaimer_style = ParagraphStyle(
            'DisclaimerStyle',
            fontSize=8,
            textColor=colors.grey,
            alignment=0,  # Sinistra
            leading=10
        )

        header_data = []
        disclaimer_text = "Documento strettamente riservato ad uso interno<br/>COEMI s.r.l."

        if os.path.exists(logo_path):
            img = Image(logo_path, width=2.5 * cm, height=2.5 * cm)
            # La cella sinistra contiene Logo + Spacer + Disclaimer
            left_cell = [img, Spacer(1, 0.3 * cm), Paragraph(disclaimer_text, disclaimer_style)]
            header_data = [[left_cell, Paragraph(f"<b>REPORT PRENOTAZIONE PDL</b><br/>Generato il: {now_str}<br/>Rif: SAF-PRN-{now.strftime('%d%m%Y')}<br/>Sistema: SafeWork-PDL v2.1.0", meta_style)]]
        else:
            header_data = [[Paragraph(f"<b>COEMI S.R.L.</b><br/><font size='6'>{disclaimer_text}</font>", styles['Normal']), Paragraph(f"<b>REPORT PRENOTAZIONE PDL</b><br/>Generato il: {now_str}<br/>Rif: SAF-PRN-{now.strftime('%d%m%Y')}<br/>Sistema: SafeWork-PDL v2.1.0", meta_style)]]

        header_table = Table(header_data, colWidths=[8 * cm, 10 * cm])
        header_table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
        ]))
        elements.append(header_table)

        # Linea di separazione
        elements.append(Spacer(1, 0.2 * cm))
        line_table = Table([['']], colWidths=[18 * cm], rowHeights=[0.1 * cm])
        line_table.setStyle(TableStyle([
            ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),
        ]))
        elements.append(line_table)
        elements.append(Spacer(1, 0.8 * cm))

        # --- 2. DASHBOARD DI RIEPILOGO ---
        eseguite = sum(1 for p in pdl_list if "successo" in str(p.stato_script).lower())
        gia_prenotate = sum(1 for p in pdl_list if "già prenotato" in str(p.stato_script).lower())
        errori = len(pdl_list) - eseguite - gia_prenotate

        dash_style_label = ParagraphStyle('DashLabel', fontSize=7, alignment=1, textColor=colors.grey)
        dash_style_value = ParagraphStyle('DashValue', fontSize=14, alignment=1, fontName='Helvetica-Bold')

        dash_data = [
            [Paragraph("TOTALE PDL", dash_style_label),
             Paragraph("ESEGUITE OGGI", dash_style_label),
             Paragraph("GIÀ PRENOTATE", dash_style_label),
             Paragraph("NON ESEGUITE", dash_style_label)],
            [Paragraph(str(len(pdl_list)), dash_style_value),
             Paragraph(str(eseguite), dash_style_value),
             Paragraph(str(gia_prenotate), dash_style_value),
             Paragraph(str(errori), dash_style_value)]
        ]

        dash_table = Table(dash_data, colWidths=[4.5 * cm, 4.5 * cm, 4.5 * cm, 4.5 * cm])
        dash_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.whitesmoke),
            ('BOX', (0, 0), (-1, -1), 0.5, colors.lightgrey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ]))
        elements.append(dash_table)
        elements.append(Spacer(1, 1 * cm))

        # --- 3. DETTAGLIO ATTIVITÀ (TABELLE DISTINTE PER AREA) ---
        elements.append(Paragraph("<b>DETTAGLIO ATTIVITÀ PER AREA</b>", ParagraphStyle('SectionTitle', fontSize=12, spaceAfter=12)))

        # Raggruppamento dei dati per Area
        areas: dict[str, list[PDLData]] = {}
        for pdl in pdl_list:
            area_name = str(pdl.area or "AREA NON SPECIFICATA").upper()
            if area_name not in areas:
                areas[area_name] = []
            areas[area_name].append(pdl)

        # Ordinamento alfabetico delle aree per un report consistente
        sorted_areas = sorted(areas.keys())

        for area_name in sorted_areas:
            # Titolo Sezione Area
            elements.append(Paragraph(f"{area_name}", ParagraphStyle('AreaHeader', fontSize=10, fontName='Helvetica-Bold', spaceBefore=10, spaceAfter=6, leftIndent=0)))

            # Intestazioni tabella per questa Area
            data = [['PdL', 'Impianto', 'Orario Prenotazione', 'Stato Esito']]

            for pdl in areas[area_name]:
                tempo_pulito = str(pdl.tempo_rimanente or "-").split('(')[0].strip()
                stato = str(pdl.stato_script)
                if len(stato) > max_status_len:
                    stato = stato[:max_status_len - 3] + "..."

                data.append([
                    Paragraph(f"<b>{pdl.pdl}</b>", styles['Normal']),
                    str(pdl.impianto),
                    tempo_pulito,
                    Paragraph(f"<i>{stato}</i>", styles['Normal'])
                ])

            # Creazione Tabella per l'Area corrente
            col_widths = [3.5 * cm, 5.5 * cm, 4.5 * cm, 4.5 * cm]
            table = Table(data, colWidths=col_widths, repeatRows=1)

            # Stile Tabella (Moderno B/N)
            current_table_style = TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.black),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 9),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                ('TOPPADDING', (0, 0), (-1, 0), 8),

                # Corpo: Linee orizzontali e Zebra
                ('GRID', (0, 0), (-1, -1), 0, colors.white),
                ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),
                ('LINEBELOW', (0, 1), (-1, -1), 0.1, colors.lightgrey),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 9),
                ('TOPPADDING', (0, 1), (-1, -1), 6),
                ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
            ])

            # Zebra striping per la tabella corrente
            for i in range(1, len(data)):
                if i % 2 == 0:
                    current_table_style.add('BACKGROUND', (0, i), (-1, i), colors.whitesmoke)

            table.setStyle(current_table_style)
            elements.append(table)
            elements.append(Spacer(1, 0.6 * cm))

        # Salvataggio con numerazione pagine dinamica
        doc.build(elements, onFirstPage=self._add_footer, onLaterPages=self._add_footer)
        logger.info(f"PDF generato con successo: {file_path}")

    def _add_footer(self, canvas: Any, doc: Any) -> None:
        """Aggiunge il piè di pagina dinamico con numerazione Pagina X di Y."""
        canvas.saveState()

        # Linea sottile sopra il footer
        canvas.setStrokeColor(colors.lightgrey)
        canvas.setLineWidth(0.5)
        canvas.line(1.5 * cm, 1.5 * cm, A4[0] - 1.5 * cm, 1.5 * cm)

        # Numerazione pagina
        page_num = canvas.getPageNumber()
        # Nota: ReportLab non conosce il totale pagine a priori senza due passaggi,
        # ma per semplicità e performance usiamo "Pagina X" o un metodo standard.
        # Per avere "X di Y" servirebbe un post-processing. Procediamo con "Pagina X".
        text = f"Pagina {page_num}"

        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.grey)
        canvas.drawRightString(A4[0] - 1.5 * cm, 1.1 * cm, text)

        canvas.restoreState()
