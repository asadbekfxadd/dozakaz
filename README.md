# Дозаказы — система управления заявками

## Запуск локально
```bash
pip install -r requirements.txt
python app.py
```
Откроется на http://localhost:5000

## Деплой на Railway
1. Зайди на railway.app
2. New Project → Deploy from GitHub repo
3. Загрузи эту папку или подключи GitHub
4. Railway автоматически запустит приложение

## Структура
- `app.py` — Flask сервер + API
- `templates/index.html` — фронтенд
- `uploads/` — загруженные Excel файлы
- `dozakaz.db` — база данных SQLite (создаётся автоматически)

## Пароли (важно)
Пароли больше не хранятся в коде. Задайте переменные в Railway:
- `SECRET_KEY` — длинная случайная строка (иначе сессии сбрасываются при каждом перезапуске).
- `USER_PASSWORDS` — JSON с паролями (или их хэшами) по логину, например `{"admin":"...","DMITRIY":"...","SKLAD":"...","ANNA_V":"...","ANNA_G":"..."}`.
- `BRANCH_PASSWORD` — общий пароль для всех филиалов (роль `user`), если у филиала нет своей записи в `USER_PASSWORDS`.
Нет пароля в окружении — войти этим пользователем нельзя. Смена пароля завершает старые сессии.
Хэш вместо пароля: `python -c "from werkzeug.security import generate_password_hash as g; print(g('пароль'))"`.

## Тесты
`pip install -r requirements-dev.txt` и `TEST_DATABASE_URL=postgresql://user:pass@host/db pytest tests` (нужен пустой PostgreSQL).
