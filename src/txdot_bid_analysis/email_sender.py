"""Send the four generated report files through Gmail."""
from email.message import EmailMessage
import os
from pathlib import Path
import smtplib
import ssl

from dotenv import load_dotenv

from .config import DEFAULT_PROJECT_ID


def send_email(project_id: str = DEFAULT_PROJECT_ID, output_dir: Path = Path('output')) -> None:
    """Load local Gmail settings, attach the reports, and send once."""
    load_dotenv(Path.cwd()/'.env', override=False)
    sender = os.getenv('EMAIL_FROM', '').strip()
    recipient = os.getenv('EMAIL_TO', '').strip()
    password = os.getenv('GMAIL_APP_PASSWORD', '').replace(' ', '').strip()
    if not sender or not recipient or not password:
        raise ValueError('Set EMAIL_FROM, EMAIL_TO and GMAIL_APP_PASSWORD in .env first.')
    if not project_id or any(c not in 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-' for c in project_id):
        raise ValueError('Invalid project ID.')

    files = [output_dir/project_id/'bid_comparison.xlsx',
             output_dir/project_id/'bid_comparison.pdf',
             output_dir/'master_unit_prices.xlsx',
             output_dir/'estimates'/'estimate.xlsx']
    message = EmailMessage()
    message['From'] = sender
    message['To'] = recipient
    message['Subject'] = f'TxDOT bid analysis — {project_id}'
    message.set_content('Hello,\n\nAttached are the bid comparisons, historical master, '
                        'and interactive estimate workbook. Estimates use historical prices.\n\nThank you.')
    for path in files:
        subtype = 'pdf' if path.suffix == '.pdf' else 'vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        message.add_attachment(path.read_bytes(), maintype='application', subtype=subtype, filename=path.name)

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=ssl.create_default_context(), timeout=30) as smtp:
            smtp.login(sender, password)
            if smtp.send_message(message, from_addr=sender, to_addrs=[recipient]):
                raise RuntimeError('Gmail did not accept the recipient.')
    except (smtplib.SMTPException, OSError):
        # Provider errors may contain credentials; do not expose them or retry automatically.
        raise RuntimeError('Email send failed or delivery is uncertain. Check Gmail Sent before retrying.') from None
