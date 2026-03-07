"""
Запуск один раз — для GitHub Actions.
Забирает отчёт из почты и рассылает в Telegram.
"""
import sys
import os
import ssl
import imaplib
import email
import re
import logging
from email.header import decode_header
from datetime import datetime
from pathlib import Path

# Настройки
IMAP_SERVER = "mail.kari.com"
IMAP_PORT   = 993
EMAIL_LOGIN = "a.kustov@kari.com"
EMAIL_PASSWORD = "a5YFpmM!"

MAILBOXES = ["INBOX", "&BD4EQgRHBDUEQgRL-"]
SENDER_EMAIL = "reports@kari.com"
SUBJECT_MUST_CONTAIN = ["часу продаж", "ростов"]
FILENAME_KEYWORDS    = ["подразделение", "часу"]

BASE_DIR     = Path(__file__).parent
DOWNLOAD_DIR = BASE_DIR / "reports"
LOG_FILE     = BASE_DIR / "run.log"

DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(message)s",
    datefmt="%d.%m.%Y %H:%M:%S",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)


def make_ssl_ctx():
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    return ctx


def decode_str(value):
    if not value:
        return ""
    parts = decode_header(value)
    result = []
    for text, charset in parts:
        if isinstance(text, bytes):
            result.append(text.decode(charset or "utf-8", errors="replace"))
        else:
            result.append(text)
    return "".join(result)


def sanitize(name):
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip()


def fetch_latest_report():
    ctx = make_ssl_ctx()
    mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT, ssl_context=ctx)
    mail.login(EMAIL_LOGIN, EMAIL_PASSWORD)

    latest_file = None

    for mailbox in MAILBOXES:
        mail.select(mailbox)
        status, data = mail.search(None, f'FROM "{SENDER_EMAIL}"')
        ids = data[0].split()[-100:]
        log.info(f"  Папка {mailbox!r}: {len(ids)} писем от {SENDER_EMAIL}")

        for num in reversed(ids):
            try:
                status, hdr_data = mail.fetch(num, "(BODY[HEADER.FIELDS (FROM SUBJECT DATE)])")
                if status != "OK":
                    continue
                hdr = email.message_from_bytes(hdr_data[0][1])
                subject = decode_str(hdr.get("Subject", ""))
                subj_lower = subject.lower()

                if not all(kw.lower() in subj_lower for kw in SUBJECT_MUST_CONTAIN):
                    continue

                log.info(f"  📧 Письмо: «{subject}»")

                status, msg_data = mail.fetch(num, "(RFC822)")
                if status != "OK":
                    continue
                msg = email.message_from_bytes(msg_data[0][1])

                for part in msg.walk():
                    filename = part.get_filename()
                    if not filename:
                        continue
                    filename = decode_str(filename)
                    fname_lower = filename.lower()
                    if not fname_lower.endswith(".xlsx"):
                        continue
                    if not any(kw.lower() in fname_lower for kw in FILENAME_KEYWORDS):
                        continue

                    ts = datetime.now().strftime("%Y%m%d_%H%M")
                    safe_name = f"{sanitize(filename.rsplit('.', 1)[0])}_{ts}.xlsx"
                    filepath = DOWNLOAD_DIR / safe_name
                    filepath.write_bytes(part.get_payload(decode=True))
                    log.info(f"  💾 Сохранён: {filepath}")
                    latest_file = str(filepath)
                    break

            except Exception as e:
                log.error(f"  Ошибка при обработке письма {num}: {e}")

            if latest_file:
                break

        if latest_file:
            break

    try:
        mail.logout()
    except Exception:
        pass

    return latest_file


if __name__ == "__main__":
    log.info("=" * 55)
    log.info("🚀 GitHub Actions — запуск отчёта")
    log.info(f"🕐  Время: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}")
    log.info("=" * 55)

    filepath = fetch_latest_report()

    if filepath:
        log.info(f"✅ Файл найден: {filepath}")
        sys.path.insert(0, str(BASE_DIR))
        from telegram_sender import main as send_report
        send_report(filepath)
        log.info("✅ Рассылка завершена!")
    else:
        log.warning("⚠️  Отчёт не найден в почте на данный момент.")
        sys.exit(0)
