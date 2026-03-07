#!/usr/bin/env python3
"""
ОТПРАВЩИК В TELEGRAM — Подразделение ЮГ
Использование: python telegram_sender.py путь_к_файлу.xlsx
"""
import requests, sys, time, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from bot_analyzer import run

BOT_TOKEN = "8604841476:AAGnYTalL6v6rFLW2qHJwTSNJ0F9kiHf8oY"

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

def send(chat_id, text, parse_mode="HTML"):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage"
    for chunk in [text[i:i+4000] for i in range(0, len(text), 4000)]:
        resp = requests.post(url, json={"chat_id": chat_id, "text": chunk, "parse_mode": parse_mode})
        print("    ✅ OK" if resp.status_code == 200 else f"    ❌ {resp.text}")
        time.sleep(0.5)

def send_file(chat_id, filepath):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendDocument"
    with open(filepath, "rb") as f:
        resp = requests.post(url, data={"chat_id": chat_id}, files={"document": f})
    print("    ✅ Файл OK" if resp.status_code == 200 else f"    ❌ {resp.text}")

def main(filepath):
    print(f"\n📂 Файл: {filepath}")
    result = run(filepath)

    print("\n📢 Общая сводка → Текущая работа РОСТОВ")
    send(CHAT_IDS["ОБЩИЙ ЧАТ ПОДРАЗДЕЛЕНИЯ"],
         result['general']['message'],
         result['general']['parse_mode'])
    send_file(CHAT_IDS["ОБЩИЙ ЧАТ ПОДРАЗДЕЛЕНИЯ"], filepath)

    print(f"\n✅ Готово! Отправлено в общий чат.")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Использование: python telegram_sender.py файл.xlsx")
        sys.exit(1)
    main(sys.argv[1])
