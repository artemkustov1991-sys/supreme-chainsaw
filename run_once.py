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
import argparse
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

BASE_DIR       = Path(__file__).parent
DOWNLOAD_DIR   = BASE_DIR / "reports"
LOG_FILE       = BASE_DIR / "run.log"
SENT_FLAGS_DIR = Path(os.environ.get("SENT_FLAGS_DIR", str(BASE_DIR / ".sent_flags")))

MOSCOW_TZ = timezone(timedelta(hours=3))

# Карта: (от_часа, до_часа_включительно) → метка времени в теле письма
# Окно на час раньше — чтобы ловить запуски в :50 предыдущего часа
SLOT_MAP = {
    (12, 14): "12:00",
    (15, 17): "15:00",
    (18, 20): "18:00",
    (22, 24): "22:00",
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
    Возвращает метку времени (например '12:00') или None если вне окна.
    Окна расширены с учётом задержки GitHub Actions до 2+ часов."""
    now_msk = datetime.now(MOSCOW_TZ)
    h = now_msk.hour
    # Широкие окна: каждый слот покрывает ~3 часа задержки
    # 22:00 покрывает переход через полночь (h=22,23,0,1)
    if 12 <= h < 15:
        return "12:00"
    if 15 <= h < 18:
        return "15:00"
    if 18 <= h < 22:
        return "18:00"
    if h >= 22 or h < 2:   # 22:00–23:59 и 00:00–01:59 МСК
        return "22:00"
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
def get_report_date(time_label):
    """Возвращает дату письма МСК. Слот 22:00 после полуночи → ищем вчера."""
    now_msk = datetime.now(MOSCOW_TZ)
    if time_label == "22:00" and now_msk.hour < 2:
        return (now_msk - timedelta(days=1)).date()
    return now_msk.date()


def fetch_latest_report(time_label):
    """Ищет НЕФЛАЖЕННОЕ письмо с нужным временным слотом.
    Сразу помечает найденное письмо флагом \\Flagged на IMAP-сервере —
    это единый замок для всех машин (GitHub Actions, Termux, ПК).
    Возвращает путь к xlsx или None."""
    today_msk = get_report_date(time_label)
    # IMAP дата для поиска (формат: DD-Mon-YYYY)
    imap_date = today_msk.strftime("%d-%b-%Y")

    ctx = make_ssl_ctx()
    mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT, ssl_context=ctx)
    mail.login(EMAIL_LOGIN, EMAIL_PASSWORD)

    latest_file = None

    for mailbox in MAILBOXES:
        mail.select(mailbox)
        # Ищем только НЕФЛАЖЕННЫЕ письма — флаг \Flagged = «уже обработано»
        status, data = mail.search(
            None,
            f'UNFLAGGED FROM "{SENDER_EMAIL}" ON "{imap_date}"'
        )
        ids = data[0].split()
        log.info(f"  Папка {mailbox!r}: {len(ids)} нефлаженных писем от {SENDER_EMAIL} за {today_msk}")

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

                # ФЛАЖИМ НЕМЕДЛЕННО — атомарный «замок» на почтовом сервере.
                # Следующий запуск (GA / Termux) не найдёт письмо в UNFLAGGED.
                try:
                    mail.store(num, '+FLAGS', '\\Flagged')
                    log.info(f"  Письмо помечено \\Flagged (дедупликация)")
                except Exception as fe:
                    log.warning(f"  Не удалось поставить \\Flagged: {fe}")

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


# ── Защита от двойной отправки ────────────────────────────────────────────────
def sent_flag_path(today, time_label):
    return SENT_FLAGS_DIR / f"{today}_{time_label.replace(':', '-')}.sent"

def already_sent(today, time_label):
    return sent_flag_path(today, time_label).exists()

def mark_sent(today, time_label):
    """Атомарно создаёт флаг — возвращает True если удалось (мы первые), False если уже занято."""
    SENT_FLAGS_DIR.mkdir(parents=True, exist_ok=True)
    try:
        sent_flag_path(today, time_label).open('x').close()
        return True
    except FileExistsError:
        return False


# ── Точка входа ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--slot", default=None, help="Временной слот, например 12:00")
    args = parser.parse_args()

    now_msk = datetime.now(MOSCOW_TZ)
    log.info("=" * 55)
    log.info(f"Запуск отчёта")
    log.info(f"Время МСК: {now_msk.strftime('%d.%m.%Y %H:%M:%S')}")

    if args.slot:
        time_label = args.slot
        log.info(f"Слот задан аргументом: {time_label}")
    else:
        time_label = get_target_slot()

    if not time_label:
        log.warning(f"Вне рабочего окна (час МСК: {now_msk.hour}). Выход.")
        sys.exit(0)

    # Дата отчёта (для 22:00 после полуночи — вчера)
    today = get_report_date(time_label).strftime("%Y-%m-%d")
    log.info(f"Дата отчёта: {today}")

    if already_sent(today, time_label):
        log.info(f"Отчёт {time_label} за {today} уже отправлен. Пропуск.")
        sys.exit(0)

    log.info(f"Ищем отчёт за слот: {time_label}")
    log.info("=" * 55)

    filepath = fetch_latest_report(time_label)

    if not filepath:
        log.warning(f"Отчёт '{time_label}' за сегодня не найден. Ничего не отправляем.")
        sys.exit(0)

    log.info(f"Файл найден: {filepath}")

    # Атомарно занимаем слот ДО отправки — защита от повторной отправки
    if not mark_sent(today, time_label):
        log.info(f"Отчёт {time_label} уже занят другим процессом. Пропуск.")
        sys.exit(0)

    sys.path.insert(0, str(BASE_DIR))
    from telegram_sender import main as send_report
    send_report(filepath)
    log.info("Рассылка завершена!")
