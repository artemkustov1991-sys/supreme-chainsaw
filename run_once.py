"""
Запуск один раз — для GitHub Actions.
Забирает отчёт из почты и рассылает в Telegram.
Логика: определяет текущий временной слот (МСК) и ищет
соответствующее письмо за сегодня.
"""
import sys
import os
import ssl
import imaplib
import email
import re
import logging
from email.header import decode_header
from email.utils import parsedate_to_datetime
from datetime import datetime, timezone, timedelta
from pathlib import Path

# ── Настройки ────────────────────────────────────────────────────────────────
IMAP_SERVER    = "mail.kari.com"
IMAP_PORT      = 993
EMAIL_LOGIN    = "a.kustov@kari.com"
EMAIL_PASSWORD = "a5YFpmM!"

MAILBOXES           = ["INBOX", "&BD4EQgRHBDUEQgRL-"]
SENDER_EMAIL        = "reports@kari.com"
SUBJECT_MUST_CONTAIN = ["часу продаж", "ростов", "для подразделени"]
FILENAME_KEYWORDS   = ["подразделение", "часу"]

BASE_DIR     = Path(__file__).parent
DOWNLOAD_DIR = BASE_DIR / "reports"
LOG_FILE     = BASE_DIR / "run.log"

MOSCOW_TZ = timezone(timedelta(hours=3))

# Карта: (от_часа, до_часа_включительно) → метка времени в теле письма
SLOT_MAP = {
    (13, 14): "12:00",
    (16, 17): "15:00",
    (19, 20): "18:00",
    (23, 24): "22:00",
}

# ── Логирование ───────────────────────────────────────────────────────────────
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


# ── Вспомогательные функции ───────────────────────────────────────────────────
def get_target_slot():
    """Определяет какой временной слот ищем по текущему времени МСК.
    Возвращает метку времени (например '12-00') или None если вне окна."""
    now_msk = datetime.now(MOSCOW_TZ)
    h = now_msk.hour
    for (h_from, h_to), label in SLOT_MAP.items():
        if h_from <= h < h_to:
            return label
    # GitHub Actions может опоздать на 30+ мин — расширяем окно
    for (h_from, h_to), label in SLOT_MAP.items():
        if h_from <= h < h_to + 1:
            return label
    return None


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


def get_email_date_msk(msg):
    """Возвращает дату письма в МСК или None."""
    try:
        date_str = msg.get("Date", "")
        dt = parsedate_to_datetime(date_str)
        return dt.astimezone(MOSCOW_TZ)
    except Exception:
        return None


def body_contains_time(msg, time_label):
    """Проверяет есть ли метка времени (например '12-00') в теле письма."""
    for part in msg.walk():
        ct = part.get_content_type()
        if ct in ("text/plain", "text/html"):
            try:
                payload = part.get_payload(decode=True)
                if payload:
                    text = payload.decode("utf-8", errors="replace")
                    if time_label in text:
                        return True
            except Exception:
                pass
    return False


# ── Основная функция ──────────────────────────────────────────────────────────
def fetch_latest_report(time_label):
    """Ищет сегодняшнее письмо с нужным временным слотом, возвращает путь к xlsx."""
    today_msk = datetime.now(MOSCOW_TZ).date()
    # IMAP дата для поиска (формат: DD-Mon-YYYY)
    imap_date = today_msk.strftime("%d-%b-%Y")

    ctx = make_ssl_ctx()
    mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT, ssl_context=ctx)
    mail.login(EMAIL_LOGIN, EMAIL_PASSWORD)

    latest_file = None

    for mailbox in MAILBOXES:
        mail.select(mailbox)
        # Ищем сегодняшние письма от нужного отправителя
        status, data = mail.search(
            None,
            f'FROM "{SENDER_EMAIL}" ON "{imap_date}"'
        )
        ids = data[0].split()
        log.info(f"  Папка {mailbox!r}: {len(ids)} писем от {SENDER_EMAIL} за {today_msk}")

        # Перебираем от новых к старым
        for num in reversed(ids):
            try:
                status, msg_data = mail.fetch(num, "(RFC822)")
                if status != "OK":
                    continue
                msg = email.message_from_bytes(msg_data[0][1])

                # Проверяем тему
                subject = decode_str(msg.get("Subject", ""))
                subj_lower = subject.lower()
                if not all(kw.lower() in subj_lower for kw in SUBJECT_MUST_CONTAIN):
                    continue

                # Проверяем дату письма — только сегодня (МСК)
                email_dt = get_email_date_msk(msg)
                if email_dt and email_dt.date() != today_msk:
                    log.info(f"  Пропуск (не сегодня): {subject} [{email_dt.date()}]")
                    continue

                # Проверяем время в теле письма
                if not body_contains_time(msg, time_label):
                    log.info(f"  Пропуск (нет '{time_label}' в теле): {subject}")
                    continue

                log.info(f"  Письмо подходит: {subject} | {email_dt}")

                # Берём только xlsx вложение
                for part in msg.walk():
                    filename = part.get_filename()
                    if not filename:
                        continue
                    filename = decode_str(filename)
                    if not filename.lower().endswith(".xlsx"):
                        continue
                    if not any(kw.lower() in filename.lower() for kw in FILENAME_KEYWORDS):
                        continue

                    ts = datetime.now(MOSCOW_TZ).strftime("%Y%m%d_%H%M")
                    safe_name = f"{sanitize(filename.rsplit('.', 1)[0])}_{ts}.xlsx"
                    filepath = DOWNLOAD_DIR / safe_name
                    filepath.write_bytes(part.get_payload(decode=True))
                    log.info(f"  Сохранён: {filepath.name}")
                    latest_file = str(filepath)
                    break

            except Exception as e:
                log.error(f"  Ошибка письма {num}: {e}")

            if latest_file:
                break

        if latest_file:
            break

    try:
        mail.logout()
    except Exception:
        pass

    return latest_file


# ── Точка входа ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    now_msk = datetime.now(MOSCOW_TZ)
    log.info("=" * 55)
    log.info(f"GitHub Actions — запуск отчёта")
    log.info(f"Время МСК: {now_msk.strftime('%d.%m.%Y %H:%M:%S')}")

    time_label = get_target_slot()

    if not time_label:
        log.warning(f"Вне рабочего окна (час МСК: {now_msk.hour}). Выход.")
        sys.exit(0)

    log.info(f"Ищем отчёт за слот: {time_label}")
    log.info("=" * 55)

    filepath = fetch_latest_report(time_label)

    if not filepath:
        log.warning(f"Отчёт '{time_label}' за сегодня не найден. Ничего не отправляем.")
        sys.exit(0)

    log.info(f"Файл найден: {filepath}")
    sys.path.insert(0, str(BASE_DIR))
    from telegram_sender import main as send_report
    send_report(filepath)
    log.info("Рассылка завершена!")
