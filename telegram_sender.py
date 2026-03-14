#!/usr/bin/env python3
"""
ОТПРАВЩИК В TELEGRAM — Подразделение ЮГ
Использование: python telegram_sender.py путь_к_файлу.xlsx
"""
import requests, sys, time, os, io
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
if hasattr(sys.stderr, 'buffer'):
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from bot_analyzer import run, load, make_praise
from report_image import generate_images
from recommendations import get_tip

BOT_TOKEN = "8604841476:AAGnYTalL6v6rFLW2qHJwTSNJ0F9kiHf8oY"

REPORTS_THREAD_ID = 2   # топик "отчеты" в группе "Текущая работа РОСТОВ"

CHAT_IDS = {
    "ОБЩИЙ ЧАТ ПОДРАЗДЕЛЕНИЯ": "-1002696361907",   # Текущая работа РОСТОВ
    "10045 Шахты Ростов":          "-1001842949200",
    "10262 Кореновск Ростов":      "-1001889416106",
    "10347 Батайск Ростов":        "-1001660993460",
    "10775 Тихорецк Ростов":       "-1001765700513",
    "10942 Новочеркасск Ростов":   "-879793202",
    "10969 Шахты Ростов":          "-854022686",
    "11193 Белая Калитва Ростов":  "-820612493",
    "11208 Азов (Ростов)":         "-1001767342727",
    "11267 Донецк Ростов":         "-1001880351055",
    "11338 Динская Ростов":        "-4931993438",
    "11497 Каменск Ростов":        "-830706279",
    "11666 Тимашевск Ростов":      "-1001879257708",
    "11667 Батайск Черноморское":  "-1001815761420",
    "11682 Новочеркасск Ростов":   "-1002028855894",
    "11694 Гуково Ростов":         "-1001868589486",
    "11840 Миллерово Ростов":      "-1001771189416",
    "11895_Новошахтинск Ростов":   "-1001809726314",
    "13034 Волгодонск Ростов":     "-806406635",
    "13125_Элиста_Ростов":         "-1002178607531",
    "13159_Сальск_Ростов":         "-1002164998662",
}

def send(chat_id, text, parse_mode="HTML", thread_id=None):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage"
    for chunk in [text[i:i+4000] for i in range(0, len(text), 4000)]:
        payload = {"chat_id": chat_id, "text": chunk, "parse_mode": parse_mode}
        if thread_id:
            payload["message_thread_id"] = thread_id
        resp = requests.post(url, json=payload)
        print("    ✅ OK" if resp.status_code == 200 else f"    ❌ {resp.text}")
        time.sleep(0.5)

def send_file(chat_id, filepath, thread_id=None):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendDocument"
    data = {"chat_id": chat_id}
    if thread_id:
        data["message_thread_id"] = thread_id
    with open(filepath, "rb") as f:
        resp = requests.post(url, data=data, files={"document": f})
    print("    ✅ Файл OK" if resp.status_code == 200 else f"    ❌ {resp.text}")

def send_photo(chat_id, png_bytes, thread_id=None):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendPhoto"
    data = {"chat_id": chat_id}
    if thread_id:
        data["message_thread_id"] = thread_id
    resp = requests.post(url, data=data, files={"photo": ("table.png", png_bytes, "image/png")})
    print("    ✅ Фото OK" if resp.status_code == 200 else f"    ❌ {resp.text}")
    time.sleep(0.5)

def main(filepath):
    print(f"\n📂 Файл: {filepath}")
    result = run(filepath)
    stores, date_str, time_str, norms, total = load(filepath)

    chat_id = CHAT_IDS["ОБЩИЙ ЧАТ ПОДРАЗДЕЛЕНИЯ"]
    tid = REPORTS_THREAD_ID

    print("\n📢 Общая сводка → топик 'отчеты'")
    send(chat_id, result['general']['message'], result['general']['parse_mode'], thread_id=tid)
    send_file(chat_id, filepath, thread_id=tid)

    print("\n📊 Генерация таблиц-рейтингов...")
    images = generate_images(stores, norms, total, date_str, time_str)
    for label, png_bytes in images:
        print(f"    → {label}")
        send_photo(chat_id, png_bytes, thread_id=tid)

    if time_str == "22:00":
        print("\n🌟 Итоги дня — похвала лучших...")
        send(chat_id, make_praise(stores, date_str), thread_id=tid)
    else:
        print("\n💡 Отправка рекомендаций...")
        send(chat_id, get_tip(), thread_id=tid)

    print(f"\n✅ Готово! Отправлено в общий чат.")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Использование: python telegram_sender.py файл.xlsx")
        sys.exit(1)
    main(sys.argv[1])
