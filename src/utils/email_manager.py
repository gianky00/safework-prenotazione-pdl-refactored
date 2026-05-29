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
        """Costruisce il corpo HTML per il report di successo."""
        rows = ""
        for pdl in pdl_list:
            stato = (pdl.stato_script or "").upper()
            # Verde per prenotazione eseguita, rosso per errori, arancione per il resto
            if "PRENOTAZIONE ESEGUITA" in stato or "OK" in stato:
                status_color = "#28a745"  # Verde Successo
            elif "ERR" in stato or "FALLITO" in stato:
                status_color = "#dc3545"  # Rosso Errore
            else:
                status_color = "#ff8c00"  # Arancione Warning/Info

            # Pulisce l'orario togliendo le parentesi
            orario_pulito = str(pdl.tempo_rimanente or "-").split('(')[0].strip()

            rows += f"""
            <tr>
                <td style="border: 1px solid #ddd; padding: 8px; white-space: nowrap;">{pdl.pdl}</td>
                <td style="border: 1px solid #ddd; padding: 8px; white-space: nowrap;">{pdl.area}</td>
                <td style="border: 1px solid #ddd; padding: 8px; white-space: nowrap;">{pdl.impianto}</td>
                <td style="border: 1px solid #ddd; padding: 8px; white-space: nowrap;">{orario_pulito}</td>
                <td style="border: 1px solid #ddd; padding: 8px; white-space: nowrap; color: {status_color};"><b>{pdl.stato_script}</b></td>
            </tr>
            """

        html = f"""
        <html>
        <body style="font-family: Arial, sans-serif;">
            <h2 style="color: #0088ff;">Report Automazione Prenotazione PDL</h2>
            <p>Di seguito il riepilogo dell'elaborazione effettuata in data {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}:</p>
            <table style="border-collapse: collapse; width: auto; min-width: 500px;">
                <tr style="background-color: #f2f2f2;">
                    <th style="border: 1px solid #ddd; padding: 8px; text-align: left; white-space: nowrap;">PdL</th>
                    <th style="border: 1px solid #ddd; padding: 8px; text-align: left; white-space: nowrap;">Area</th>
                    <th style="border: 1px solid #ddd; padding: 8px; text-align: left; white-space: nowrap;">Impianto</th>
                    <th style="border: 1px solid #ddd; padding: 8px; text-align: left; white-space: nowrap;">Orario prenotazione</th>
                    <th style="border: 1px solid #ddd; padding: 8px; text-align: left; white-space: nowrap;">Esito</th>
                </tr>
                {rows}
            </table>
            <p><i>Messaggio generato automaticamente dal sistema.</i></p>
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
