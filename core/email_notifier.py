from __future__ import annotations

import socket
import smtplib
import ssl
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders
from pathlib import Path

from config.settings import SmtpConfig
from core.logger import _now, get_logger


class EmailNotifier:
    def __init__(self, config: SmtpConfig, task_name: str):
        self.config = config
        self.task_name = task_name
        self.log = get_logger(task_name)

    def send(self, subject: str, html_body: str, attachment: Path | None = None) -> bool:
        if not self.config.enabled:
            self.log.info("Notificacion por email desactivada")
            return False

        try:
            msg = MIMEMultipart()
            msg["From"] = self.config.sender_email
            msg["To"] = ", ".join(self.config.recipient_emails)
            msg["Subject"] = subject
            msg.attach(MIMEText(html_body, "html", "utf-8"))

            if attachment and attachment.exists():
                with open(attachment, "rb") as f:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f"attachment; filename= {attachment.name}")
                msg.attach(part)
                self.log.info("Log adjuntado: %s", attachment.name)

            self._send_with_fallback(msg)
            self.log.info("Email enviado exitosamente a: %s", ", ".join(self.config.recipient_emails))
            return True

        except Exception as e:
            self.log.error("Error enviando email: %s", e)
            return False

    def _send_with_fallback(self, msg: MIMEMultipart) -> None:
        try:
            ctx = ssl.create_default_context()
            with smtplib.SMTP_SSL(self.config.server, 465, context=ctx, timeout=30) as server:
                server.login(self.config.sender_email, self.config.sender_password)
                server.send_message(msg)
            return
        except Exception:
            pass

        with smtplib.SMTP(self.config.server, self.config.port, timeout=30) as server:
            server.starttls()
            server.login(self.config.sender_email, self.config.sender_password)
            server.send_message(msg)

    def notify_success(self, detail: str, attachment: Path | None = None) -> bool:
        fecha = _now().strftime("%d/%m/%Y %H:%M:%S")  # hora Colombia
        subject = f"[OK] {self.task_name} - Completado {fecha}"
        html = f"""
        <html><body style="font-family:Arial,sans-serif;">
        <h2 style="color:#2e7d32;">✅ {self.task_name} completado</h2>
        <div style="background:#e8f5e9;padding:15px;border-left:4px solid #4caf50;margin:20px 0;">
            <p><strong>Fecha:</strong> {fecha}</p>
            <p><strong>Servidor:</strong> {socket.gethostname()}</p>
            <p><strong>Detalle:</strong> {detail}</p>
        </div>
        <hr><p style="color:#666;font-size:12px;"><em>Mensaje automatico - {self.task_name}</em></p>
        </body></html>
        """
        return self.send(subject, html, attachment)

    def notify_failure(self, error: str, attachment: Path | None = None) -> bool:
        fecha = _now().strftime("%d/%m/%Y %H:%M:%S")  # hora Colombia
        subject = f"[ERROR] {self.task_name} - Error {fecha}"
        html = f"""
        <html><body style="font-family:Arial,sans-serif;">
        <h2 style="color:#d32f2f;">❌ Error en {self.task_name}</h2>
        <div style="background:#ffebee;padding:15px;border-left:4px solid #f44336;margin:20px 0;">
            <p><strong>Fecha:</strong> {fecha}</p>
            <p><strong>Servidor:</strong> {socket.gethostname()}</p>
            <p><strong>Error:</strong></p>
            <pre style="white-space:pre-wrap;">{error}</pre>
        </div>
        <hr><p style="color:#666;font-size:12px;"><em>Mensaje automatico - {self.task_name}</em></p>
        </body></html>
        """
        return self.send(subject, html, attachment)
