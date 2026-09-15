# Развёртывание в Railway с iPhone

1. Создайте бота через @BotFather и сохраните токен.
2. Создайте бесплатный аккаунт GitHub.
3. Создайте новый приватный репозиторий, например `mossvet-telegram-bot`.
4. Загрузите в корень репозитория все файлы из этого проекта. Файл `.env` НЕ загружайте.
5. Откройте Railway и войдите через GitHub.
6. New Project → Deploy from GitHub Repo → выберите `mossvet-telegram-bot`.
7. В Variables добавьте:
   - `TELEGRAM_BOT_TOKEN` = токен BotFather
   - `DATA_ROOT` = `/data`
8. В Railway добавьте Volume и смонтируйте его в `/data`.
9. Команда запуска уже задана: `python bot.py`.
10. После Deploy откройте Telegram и нажмите Start у вашего бота.

## Важно
Токен BotFather храните только в Railway Variables. Не вставляйте его в GitHub.
