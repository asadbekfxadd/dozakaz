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
