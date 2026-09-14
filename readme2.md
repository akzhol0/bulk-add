## Сначала проверь, что в `config.json` указано:
"courseName": "Адам және жануарлар физиологиясы"

## Безопасный dry-run для Astana IT University:
npm run add -- --slug=university-26-y4lco --dry-run

Ожидаемый результат:
dry_run_ok

Это означает, что точная карточка курса типа `COURSE` найдена, но ничего не добавлено.

## Добавление в одну программу
npm run add -- --slug=university-26-y4lco

Это реально добавит курс только в Astana IT University.

## Проверка после добавления:

npm run add -- --slug=university-26-y4lco --dry-run

Ожидаемый результат:

already_exists

## Основные команды — заметка

Первоначальная установка:

Вход в Coursera — обычно один раз:

npm run login

## Проверка одной программы:

npm run add -- --slug=university-26-y4lco --dry-run

## Реальное добавление в одну программу:

npm run add -- --slug=university-26-y4lco

Проверка всех 100 без добавления:

npm run add -- --dry-run

## Реальное добавление во все 100:

npm run add

Другой курс без изменения `config.json`:

npm run add -- --slug=university-26-y4lco --course="ТОЧНОЕ НАЗВАНИЕ" --dry-run

Реальное добавление другого курса:

npm run add -- --slug=university-26-y4lco --course="ТОЧНОЕ НАЗВАНИЕ"

Важно: команда `npm run add` без `--slug` и без `--dry-run` запускает реальное добавление во все 100 программ. После каждого запуска смотри отчёт в папке `reports`.