# Personal Telegram Assistant + Google Calendar + Google Sheets

Що вміє ця версія:
- приймає текстові повідомлення
- приймає голосові та розшифровує їх через Gemini
- створює події в Google Calendar
- показує події на сьогодні / завтра / тиждень
- зберігає нотатки в Google Sheets
- веде журнал взаємодій у Google Sheets
- шукає актуальну інформацію через Google Search grounding у Gemini
- аналізує фото, які ти надсилаєш боту

## Важливо
- старий лист `Applications` не чіпається
- бот створює нові листи в тій самій таблиці:
  - `AssistantLog`
  - `AssistantNotes`
- календар використовується через `CALENDAR_ID`

## Що вже закладено в конфіг
- BOT_TOKEN уже вставлено
- GEMINI_API_KEY уже вставлено
- GOOGLE_SHEET_ID уже вставлено
- CALENDAR_ID уже вставлено як `techno.perspektiva@gmail.com`
- KV namespace уже вказано

## Що має бути в Cloudflare
Секрет:
- `GOOGLE_SERVICE_ACCOUNT_JSON`

Змінні:
- `BOT_TOKEN`
- `GEMINI_API_KEY`
- `GOOGLE_SHEET_ID`
- `WORKSHEET_NAME`
- `NOTES_WORKSHEET_NAME`
- `CALENDAR_ID`
- `TIMEZONE`
- `SETUP_KEY`
- `TELEGRAM_WEBHOOK_SECRET`
- `STATE_KV`

## Як задеплоїти
1. Завантаж цей пакет у свій репозиторій або Cloudflare Workers.
2. Переконайся, що секрет `GOOGLE_SERVICE_ACCOUNT_JSON` лишився в Worker.
3. Задеплой Worker.
4. Відкрий:
   `/setup?key=ТВОЄ_ЗНАЧЕННЯ_SETUP_KEY`
5. Напиши боту `/start`.

## Приклади фраз
- Нагадай сьогодні о 19:00 тренування
- Запиши зустріч завтра о 14:30 з Сергієм
- Занотуй: купити кабель і автомат
- Що у мене сьогодні по календарю
- Знайди мені хороший спортзал біля Позняків
- Проаналізуй це фото

## Команди
- `/start`
- `/help`
- `/today`
- `/tomorrow`
- `/week`
- `/events`
- `/note текст`
- `/reset`

## Що зберігається в Google Sheets
### AssistantLog
Журнал усіх запитів:
- created_at
- chat_id
- user_id
- username
- input_type
- original_text
- normalized_text
- action
- result_summary
- calendar_event_id
- status

### AssistantNotes
Збережені нотатки:
- created_at
- chat_id
- user_id
- username
- note_text
- source

## Рекомендація по безпеці
Після того як усе запрацює, краще винести `BOT_TOKEN`, `GEMINI_API_KEY`, `SETUP_KEY` і `TELEGRAM_WEBHOOK_SECRET` із `wrangler.jsonc` у Cloudflare Secrets.
