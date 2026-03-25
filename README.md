# HSE Dashboard

Пайплайн для сбора данных по приёмной кампании, расчёта метрик и отправки результата в Google Sheets.

## Запуск

```bash
python src/main.py
```

или

```bash
python -m src
```

## Структура данных

- `data/raw` - исходные выгрузки без ручной правки.
- `data/processed` - нормализованные промежуточные таблицы.
- `data/archive` - архивные версии и исторические снимки.
- `data/dashboards` - итоговые Excel-выгрузки.
- `templates` - эталонные шаблоны, маппинги и справочники.

Подробный контракт входных данных описан в [docs/data_contracts.md](docs/data_contracts.md).

## Зависимости

```bash
pip install -r requirements.txt
```

