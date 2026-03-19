#!/data/data/com.termux/files/usr/bin/bash
# ──────────────────────────────────────────────────────────────────────────────
# Установка и настройка авто-отправки отчётов КАРИ через Termux
# Запускать ОДИН РАЗ: bash termux_setup.sh
# ──────────────────────────────────────────────────────────────────────────────
set -e

REPO_URL="https://github.com/artemkustov1991-sys/supreme-chainsaw.git"
REPO_DIR="$HOME/kari"
LOG_FILE="$REPO_DIR/termux.log"
CRON_INTERVAL=10   # минут между проверками

echo "=== Установка пакетов ==="
pkg install -y python git cronie termux-api 2>/dev/null || true

echo "=== Клонирование / обновление репозитория ==="
if [ -d "$REPO_DIR/.git" ]; then
    cd "$REPO_DIR" && git pull --ff-only
else
    git clone "$REPO_URL" "$REPO_DIR"
fi

echo "=== Установка Python-зависимостей ==="
cd "$REPO_DIR"
pip install --quiet pandas openpyxl requests matplotlib Pillow

echo "=== Настройка cron (каждые $CRON_INTERVAL мин) ==="
# Строим строку cron — каждые 10 мин
CRON_LINE="*/$CRON_INTERVAL * * * * cd $REPO_DIR && git pull -q --ff-only 2>/dev/null; python run_once.py >> $LOG_FILE 2>&1"

# Добавляем только если ещё нет
( crontab -l 2>/dev/null | grep -v "run_once.py" ; echo "$CRON_LINE" ) | crontab -

echo "=== Запуск cron-демона ==="
pkill crond 2>/dev/null || true
crond

echo "=== Автозапуск при включении телефона ==="
mkdir -p "$HOME/.termux/boot"
cat > "$HOME/.termux/boot/kari-crond.sh" << 'BOOT'
#!/data/data/com.termux/files/usr/bin/sh
# Запуск cron при загрузке телефона
crond
BOOT
chmod +x "$HOME/.termux/boot/kari-crond.sh"

echo ""
echo "✓ Установка завершена!"
echo "  Проверки каждые $CRON_INTERVAL минут."
echo "  Логи: $LOG_FILE"
echo ""
echo "  ВАЖНО: Отключи экономию батареи для Termux:"
echo "  Настройки → Приложения → Termux → Батарея → Без ограничений"
echo ""
echo "  Проверить что cron работает:"
echo "    crontab -l"
echo "    ps aux | grep crond"
