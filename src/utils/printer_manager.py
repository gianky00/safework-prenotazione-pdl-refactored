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
            self._generate_pdf(pdl_list, file_path, project_root)
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
                win32api.ShellExecute(0, "open", file_path, None, ".", 1)
                return True

        except Exception as e:
            logger.error(f"Errore durante l'archiviazione/stampa PDF: {e}")
            return False

    def _generate_pdf(self, pdl_list: list[PDLData], file_path: str, project_root: str) -> None:
        """Costruisce il layout del PDF professionale in bianco e nero."""
        doc = SimpleDocTemplate(
            file_path,
            pagesize=A4,
            rightMargin=1.5*cm,
            leftMargin=1.5*cm,
            topMargin=1.2*cm,
            bottomMargin=1.2*cm
        )

        elements = []
        styles = getSampleStyleSheet()

        # --- 1. HEADER CON LOGO ---
        logo_path = os.path.join(project_root, "assets", "logo coemi.png")
        
        # Metadata del documento
        now_str = datetime.now().strftime("%d/%m/%Y %H:%M")
        
        meta_style = ParagraphStyle(
            'MetaStyle',
            fontSize=8,
            textColor=colors.grey,
            alignment=2 # Destra
        )
        
        header_data = []
        if os.path.exists(logo_path):
            img = Image(logo_path, width=2.5*cm, height=2.5*cm)
            header_data = [[img, Paragraph(f"<b>REPORT PRENOTAZIONE PDL</b><br/>Generato il: {now_str}<br/>Rif: SAF-PRN-{datetime.now().strftime('%y%m%d')}", meta_style)]]
        else:
            header_data = [[Paragraph("<b>COEMI S.R.L.</b>", styles['Normal']), Paragraph(f"<b>REPORT PRENOTAZIONE PDL</b><br/>Generato il: {now_str}", meta_style)]]
            
        header_table = Table(header_data, colWidths=[6*cm, 12*cm])
        header_table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
        ]))
        elements.append(header_table)
        
        # Linea di separazione
        elements.append(Spacer(1, 0.2*cm))
        line_table = Table([['']], colWidths=[18*cm], rowHeights=[0.1*cm])
        line_table.setStyle(TableStyle([
            ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),
        ]))
        elements.append(line_table)
        elements.append(Spacer(1, 0.8*cm))

        # --- 2. DASHBOARD DI RIEPILOGO ---
        successi = sum(1 for p in pdl_list if "successo" in str(p.stato_script).lower())
        errori = len(pdl_list) - successi
        
        dash_style_label = ParagraphStyle('DashLabel', fontSize=9, alignment=1, textColor=colors.grey)
        dash_style_value = ParagraphStyle('DashValue', fontSize=16, alignment=1, fontName='Helvetica-Bold')
        
        dash_data = [
            [Paragraph("TOTALE PDL", dash_style_label), Paragraph("SUCCESSI", dash_style_label), Paragraph("ERRORI / PENDING", dash_style_label)],
            [Paragraph(str(len(pdl_list)), dash_style_value), Paragraph(str(successi), dash_style_value), Paragraph(str(errori), dash_style_value)]
        ]
        
        dash_table = Table(dash_data, colWidths=[6*cm, 6*cm, 6*cm])
        dash_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.whitesmoke),
            ('BOX', (0, 0), (-1, -1), 0.5, colors.lightgrey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ]))
        elements.append(dash_table)
        elements.append(Spacer(1, 1*cm))

        # --- 3. TABELLA DETTAGLIATA (RAGGRUPPATA PER AREA) ---
        elements.append(Paragraph("<b>DETTAGLIO ATTIVITÀ PER AREA</b>", ParagraphStyle('SectionTitle', fontSize=10, spaceAfter=8)))
        
        # Ordinamento dei dati per Area per garantire il raggruppamento corretto
        pdl_list_sorted = sorted(pdl_list, key=lambda x: str(x.area))
        
        data = [['PdL', 'Impianto', 'Orario', 'Stato Esito']]
        current_area = None
        
        # Stili per la tabella
        table_style = TableStyle([
            # Intestazione Generale: Sfondo nero, testo bianco
            ('BACKGROUND', (0, 0), (-1, 0), colors.black),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
            ('TOPPADDING', (0, 0), (-1, 0), 10),
            ('GRID', (0, 0), (-1, -1), 0, colors.white), 
        ])

        row_index = 1
        for pdl in pdl_list_sorted:
            area_pdl = str(pdl.area or "AREA NON SPECIFICATA")
            
            # Se l'area cambia, aggiungiamo una riga di separazione/intestazione area
            if area_pdl != current_area:
                current_area = area_pdl
                # Riga di intestazione Area (Colspan su tutte le colonne)
                data.append([Paragraph(f"<b>ZONA: {current_area}</b>", styles['Normal']), "", "", ""])
                table_style.add('BACKGROUND', (0, row_index), (-1, row_index), colors.lightgrey)
                table_style.add('SPAN', (0, row_index), (-1, row_index))
                table_style.add('ALIGN', (0, row_index), (-1, row_index), 'LEFT')
                table_style.add('TOPPADDING', (0, row_index), (-1, row_index), 6)
                table_style.add('BOTTOMPADDING', (0, row_index), (-1, row_index), 6)
                row_index += 1

            tempo_pulito = str(pdl.tempo_rimanente or "-").split('(')[0].strip()
            stato = str(pdl.stato_script)
            if len(stato) > 25:
                stato = stato[:22] + "..."

            data.append([
                Paragraph(f"<b>{pdl.pdl}</b>", styles['Normal']),
                str(pdl.impianto),
                tempo_pulito,
                Paragraph(f"<i>{stato}</i>", styles['Normal'])
            ])
            
            # Stile righe dati
            table_style.add('LINEBELOW', (0, row_index), (-1, row_index), 0.2, colors.lightgrey)
            if row_index % 2 == 0:
                table_style.add('BACKGROUND', (0, row_index), (-1, row_index), colors.whitesmoke)
            
            row_index += 1

        col_widths = [4*cm, 5*cm, 4.5*cm, 4.5*cm]
        table = Table(data, colWidths=col_widths, repeatRows=1)

        # Stile finale righe dati
        table_style.add('FONTNAME', (0, 1), (-1, -1), 'Helvetica')
        table_style.add('FONTSIZE', (0, 1), (-1, -1), 9)
        table_style.add('VALIGN', (0, 0), (-1, -1), 'MIDDLE')
        table_style.add('ALIGN', (0, 1), (-1, -1), 'CENTER')
        
        table.setStyle(table_style)
        elements.append(table)

        # --- 4. FOOTER ---
        elements.append(Spacer(1, 1.5*cm))
        footer_data = [
            [Paragraph("Documento strettamente riservato ad uso interno - COEMI S.r.l.", ParagraphStyle('F1', fontSize=7, textColor=colors.grey)), 
             Paragraph("Sistema: SafeWork-PDL v2.1.0 | Pagina 1", ParagraphStyle('F2', fontSize=7, alignment=2, textColor=colors.grey))]
        ]
        footer_table = Table(footer_data, colWidths=[11*cm, 7*cm])
        elements.append(footer_table)

        # Salvataggio
        doc.build(elements)
        logger.info(f"PDF generato con successo: {file_path}")
