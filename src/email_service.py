"""STEP 13-14: dummy emails (.eml) with the investor's PDF attached.

Nothing is ever sent by default: config.DRY_RUN is True and send_email_outlook() refuses to run
unless DRY_RUN is turned off AND enabled=True is passed explicitly.
"""
import logging
from datetime import datetime
from email.message import EmailMessage
from email.utils import format_datetime
from pathlib import Path

import config

log = logging.getLogger("drawdown")


def recipient_address(rec):
    """DEMO_MODE -> clearly labelled dummy address (never a real person). Otherwise the Working email."""
    if config.DEMO_MODE:
        return f"{rec['folio_no'].lower()}.demo@{config.DEMO_EMAIL_DOMAIN}"
    if not rec["email"]:
        raise ValueError("Investor email is blank in Working and DEMO_MODE is off")
    return rec["email"]


def create_dummy_email(rec, pdf_path, out_path):
    """Write <folio>_<label>_Drawdown_Email.eml with the notice PDF attached. Returns the recipient."""
    pdf_path, out_path = Path(pdf_path), Path(out_path)
    to_addr = recipient_address(rec)

    msg = EmailMessage()
    msg["From"] = f"{config.CLIENT_NAME} <{config.DEMO_SENDER}>"
    msg["To"] = to_addr
    msg["Subject"] = config.EMAIL_SUBJECT.format(folio=rec["folio_no"])
    msg["Date"] = format_datetime(datetime.now().astimezone())
    msg["X-Demo-Mode"] = "true - dummy email, not sent"
    msg.set_content(config.EMAIL_BODY)
    msg.add_attachment(pdf_path.read_bytes(), maintype="application", subtype="pdf", filename=pdf_path.name)

    out_path.write_bytes(bytes(msg))
    log.info("Email created: %s -> %s", out_path.name, to_addr)
    return to_addr


def send_email_outlook(to_addr, subject, body, attachment, enabled=False):
    """Optional real sending through desktop Outlook (Windows + pywin32). Disabled unless explicitly enabled."""
    if config.DRY_RUN or not enabled:
        log.warning("DRY_RUN is on - NOT sending email to %s", to_addr)
        return False
    import win32com.client                                   # imported only when really needed
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To, mail.Subject, mail.Body = to_addr, subject, body
    mail.Attachments.Add(str(Path(attachment).resolve()))
    mail.Send()
    log.info("Email SENT to %s", to_addr)
    return True
