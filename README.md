# Dnipro-M scraper

Скрипт собирает данные по списку артикулов с сайта Dnipro-M и сохраняет результат в Excel.

## Вход

Поддерживаются файлы:

- `.txt` — один артикул на строку
- `.csv` — первый столбец содержит артикулы
- `.xlsx` — первый лист, первый столбец содержит артикулы

Повторяющиеся артикулы автоматически удаляются перед обработкой (порядок сохраняется).

## Выход

В Excel сохраняются поля:

- `Šifra proizvođača`
- `Naziv proizvoda`
- `Opis proizvoda`
- `Karakteristike`
- `Slike 1`
- `Slike 2`
- `Slike 3`

## Запуск

```bash
python scraper.py --input articles.txt --output dnipro_m_products.xlsx
```

По умолчанию скрипт использует только кэш в папке `.cache` и не пересобирает индекс страниц.

Для проблемных товаров скрипт автоматически делает повторные попытки загрузки и fallback-поиск URL.
Если артикул не найден в sitemap, скрипт дополнительно использует официальный search endpoint сайта и может найти товар по коду даже при отсутствии записи в индексе.

Если нужно явно пересобрать индекс страниц, используй:

```bash
python scraper.py --reindex
```

Если нужен явный режим только по кэшу, можно оставить `--cache-only`, но это уже не обязательно.

В cache-only режиме скрипт читает только `.cache/article_url_index.json`.

## Перевод через OpenAI

По умолчанию скрипт переводит `Naziv proizvoda`, `Opis proizvoda` и `Karakteristike` в черногорскую латиницу через OpenAI API.

Нужно задать переменную окружения:

PowerShell:

```powershell
$env:OPENAI_API_KEY="your_key_here"
```

cmd.exe:

```bash
set OPENAI_API_KEY=your_key_here
```

Опционально можно указать модель через `OPENAI_MODEL` или параметр `--translate-model`.

Если нужно сделать локальный экспорт без перевода, запусти:

```bash
python scraper.py --no-translate
```

Для перевода через OpenAI не используй `--no-translate`, а перед запуском задай `OPENAI_API_KEY`.

Запуск с переводом и ключом из txt фала с артикулами
$env:OPENAI_API_KEY = [Environment]::GetEnvironmentVariable("OPENAI_API_KEY","User")
>> python scraper.py --cache-only --input articles.txt --output dnipro_m_products.xlsx