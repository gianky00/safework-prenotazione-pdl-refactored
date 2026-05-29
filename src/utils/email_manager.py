"""Modulo per la gestione dell'invio di report via email tramite Outlook."""

from datetime import datetime
from pathlib import Path

import pythoncom
import win32com.client
from loguru import logger

from ..models import PDLData


class EmailManager:
    """Gestisce la creazione e l'invio di email tramite l'applicazione Outlook desktop."""

    DEFAULT_RECIPIENT: str = "gianky.allegretti@gmail.com"

    def __init__(self, recipient: str | None = None) -> None:
        """Inizializza il manager delle email."""
        self.recipient = recipient or self.DEFAULT_RECIPIENT

    def build_pdl_report_html(self, pdl_list: list[PDLData]) -> str:
        """Costruisce il corpo HTML per il report, allineato allo stile del PDF professionale."""
        now = datetime.now()
        now_str = now.strftime("%d/%m/%Y %H:%M")
        rif_str = f"SAF-PRN-{now.strftime('%d%m%Y')}"
        version_str = "SafeWork-PDL v2.1.0"

        # --- LOGICA CONTEGGI (Opzione B) ---
        eseguite = sum(1 for p in pdl_list if "successo" in str(p.stato_script).lower())
        gia_prenotate = sum(1 for p in pdl_list if "già prenotato" in str(p.stato_script).lower())
        errori = len(pdl_list) - eseguite - gia_prenotate

        # --- RAGGRUPPAMENTO PER AREA ---
        areas = {}
        for pdl in pdl_list:
            area_name = str(pdl.area or "AREA NON SPECIFICATA").upper()
            if area_name not in areas:
                areas[area_name] = []
            areas[area_name].append(pdl)
        sorted_areas = sorted(areas.keys())

        # --- COSTRUZIONE TABELLE PER AREA ---
        tables_html = ""
        for area_name in sorted_areas:
            rows_html = ""
            for i, pdl in enumerate(areas[area_name]):
                bg_color = "#f9f9f9" if i % 2 != 0 else "#ffffff"
                orario_pulito = str(pdl.tempo_rimanente or "-").split('(')[0].strip()
                stato = str(pdl.stato_script)
                if len(stato) > 30: stato = stato[:27] + "..."

                rows_html += f"""
                <tr style="background-color: {bg_color}; text-align: center;">
                    <td style="padding: 10px; border-bottom: 1px solid #eee;"><b>{pdl.pdl}</b></td>
                    <td style="padding: 10px; border-bottom: 1px solid #eee;">{pdl.impianto}</td>
                    <td style="padding: 10px; border-bottom: 1px solid #eee;">{orario_pulito}</td>
                    <td style="padding: 10px; border-bottom: 1px solid #eee;"><i>{stato}</i></td>
                </tr>
                """

            tables_html += f"""
            <div style="margin-top: 20px;">
                <h4 style="margin: 0 0 10px 0; color: #333; font-size: 14px;">{area_name}</h4>
                <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
                    <thead>
                        <tr style="background-color: #000000; color: #ffffff;">
                            <th style="padding: 10px; border: 1px solid #000;">PdL</th>
                            <th style="padding: 10px; border: 1px solid #000;">Impianto</th>
                            <th style="padding: 10px; border: 1px solid #000;">Orario Prenotazione</th>
                            <th style="padding: 10px; border: 1px solid #000;">Stato Esito</th>
                        </tr>
                    </thead>
                    <tbody>
                        {rows_html}
                    </tbody>
                </table>
            </div>
            """

        # --- HTML COMPLETO ---
        html = f"""
        <html>
        <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333; line-height: 1.6; margin: 20px;">
            
            <!-- HEADER -->
            <table style="width: 100%; border-bottom: 2px solid #000; padding-bottom: 10px; margin-bottom: 20px;">
                <tr>
                    <td style="width: 50%; vertical-align: middle;">
                        <h2 style="margin: 0; color: #000; font-size: 24px;">COEMI S.R.L.</h2>
                        <p style="margin: 5px 0 0 0; color: #888; font-size: 10px;">
                            Documento strettamente riservato ad uso interno<br/>
                            COEMI s.r.l.
                        </p>
                    </td>
                    <td style="width: 50%; text-align: right; vertical-align: middle; color: #888; font-size: 11px;">
                        <b style="color: #333; font-size: 13px;">REPORT PRENOTAZIONE PDL</b><br/>
                        Generato il: {now_str}<br/>
                        Rif: {rif_str}<br/>
                        Sistema: {version_str}
                    </td>
                </tr>
            </table>

            <!-- DASHBOARD -->
            <table style="width: 100%; border-collapse: collapse; margin-bottom: 30px; text-align: center;">
                <tr>
                    <td style="width: 25%; padding: 15px; background-color: #f4f4f4; border: 1px solid #ddd;">
                        <div style="font-size: 10px; color: #888; text-transform: uppercase;">Totale PDL</div>
                        <div style="font-size: 20px; font-weight: bold; color: #000;">{len(pdl_list)}</div>
                    </td>
                    <td style="width: 25%; padding: 15px; background-color: #f4f4f4; border: 1px solid #ddd;">
                        <div style="font-size: 10px; color: #888; text-transform: uppercase;">Eseguite Oggi</div>
                        <div style="font-size: 20px; font-weight: bold; color: #000;">{eseguite}</div>
                    </td>
                    <td style="width: 25%; padding: 15px; background-color: #f4f4f4; border: 1px solid #ddd;">
                        <div style="font-size: 10px; color: #888; text-transform: uppercase;">Già Prenotate</div>
                        <div style="font-size: 20px; font-weight: bold; color: #000;">{gia_prenotate}</div>
                    </td>
                    <td style="width: 25%; padding: 15px; background-color: #f4f4f4; border: 1px solid #ddd;">
                        <div style="font-size: 10px; color: #888; text-transform: uppercase;">Non Eseguite</div>
                        <div style="font-size: 20px; font-weight: bold; color: #000;">{errori}</div>
                    </td>
                </tr>
            </table>

            <!-- TITOLO SEZIONE -->
            <h3 style="border-left: 4px solid #000; padding-left: 10px; margin-bottom: 20px; font-size: 16px;">DETTAGLIO ATTIVITÀ PER AREA</h3>

            <!-- TABELLE DISTINTE -->
            {tables_html}

            <!-- FOOTER -->
            <div style="margin-top: 40px; padding-top: 10px; border-top: 1px solid #eee; font-size: 10px; color: #aaa; text-align: right;">
                Pagina 1 di 1
            </div>

        </body>
        </html>
        """
        return html

    def build_error_report_html(self, error_message: str) -> str:
        """Costruisce il corpo HTML per il report di errore critico."""
        html = f"""
        <html>
        <body style="font-family: Arial, sans-serif;">
            <h2 style="color: #ff3333;">ERRORE CRITICO - Automazione PDL</h2>
            <p>Il processo si è interrotto bruscamente a causa del seguente errore:</p>
            <div style="background-color: #fff0f0; border: 1px solid #ff3333; padding: 15px; border-radius: 5px;">
                <pre style="white-space: pre-wrap;">{error_message}</pre>
            </div>
            <p>Controllare i log allegati per maggiori dettagli.</p>
        </body>
        </html>
        """
        return html

    def send_report(
        self,
        subject: str,
        body_html: str,
        attachment_path: Path | None = None,
        display: bool = False
    ) -> bool:
        """Invia l'email tramite Outlook."""
        try:
            pythoncom.CoInitialize()
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)  # 0: olMailItem
            mail.To = self.recipient
            mail.Subject = subject
            mail.HTMLBody = body_html

            if attachment_path and attachment_path.exists():
                mail.Attachments.Add(str(attachment_path.absolute()))

            if display:
                mail.Display()
            else:
                mail.Send()
                logger.info(f"Email inviata con successo a {self.recipient}")
            return True
        except Exception as e:
            logger.error(f"Errore durante l'invio dell'email: {e}")
            return False
        finally:
            pythoncom.CoUninitialize()
