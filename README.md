# TZ Builder AI

Веб-застосунок для генерації готового DOCX технічного завдання по брифу.

## Що робить
- Аналізує текст брифу або файл з даними.
- Формує дані за фіксованою структурою (`data/output_schema.json`).
- Очищає порожні/службові артефакти в масивах перед рендером шаблону.
- Автоматично рахує `pages_count` на основі `site_sections` (з урахуванням каталогу для e-commerce).
- Рендерить фінальний документ ТЗ у форматі `.docx` через шаблон `data/template.docx`.
- Підсвічує слово `уточнити` жовтим у готовому документі.

## Запуск
1. Створіть віртуальне середовище і встановіть залежності:

```bash
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements.txt
```

2. Додайте ключ у `.env`:

```bash
cp .env.example .env
# відкрийте .env і вставте OPENAI_API_KEY
```

Додатково можна керувати таймаутом і ретраями:

```env
OPENAI_TIMEOUT_SECONDS=90
OPENAI_MAX_RETRIES=0
```

3. Запустіть сервер:

```bash
.venv/bin/python app.py
```

4. Відкрийте:

```text
http://127.0.0.1:8000
```

## Примітки
- Якщо `OPENAI_API_KEY` відсутній або API недоступне, вмикається fallback-режим: система все одно сформує DOCX, заповнивши поля за базовими правилами та евристиками.
- За замовчуванням опис плейсхолдерів береться з `data/description_merged.json`.
- Шаблон ТЗ, який використовується завжди: `data/template.docx`.
- Ви можете завантажити свій `description_merged.json` через UI для перевизначення.
- Мінімальна технічна правка для перевірки процесу пушу внесена через Codex.

## Структура
- `app.py` - Flask API + генерація + рендер DOCX.
- `web/` - UI (HTML/CSS/JS).
- `data/output_schema.json` - схема полів.
- `data/description_merged.json` - опис полів.
- `data/output_template.json` - базові значення.
- `data/template.docx` - основний шаблон документа.
