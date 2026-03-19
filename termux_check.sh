#!/data/data/com.termux/files/usr/bin/bash
# Быстрая проверка статуса и ручной запуск
REPO_DIR="$HOME/kari"
LOG_FILE="$REPO_DIR/termux.log"

echo "=== Статус cron ==="
ps aux | grep crond | grep -v grep || echo "crond НЕ запущен! Запусти: crond"

echo ""
echo "=== Расписание ==="
crontab -l 2>/dev/null || echo "Нет задач. Запусти termux_setup.sh"

echo ""
echo "=== Последние 20 строк лога ==="
tail -20 "$LOG_FILE" 2>/dev/null || echo "Лог пуст"
