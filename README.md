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


## Сбор сделок из Битрикс24

Для выгрузки всех сделок воронки «Поступление 360» из коробочного Битрикс24 используйте read-only helper:

```python
from bitrix import collect_deals_dataframe

portal_360_deals = collect_deals_dataframe()
```

Функция обращается к вебхуку `https://bx.hse.ru/rest/1/testtest/`, определяет ID воронки через `crm.category.list`, затем забирает сделки через `crm.deal.list` и пакетирует страницы методом `batch` до 50 read-only подзапросов за раз. Клиент ограничивает частоту HTTP-вызовов двумя запросами в секунду и повторяет запросы при `QUERY_LIMIT_EXCEEDED`/HTTP 503.

Если ID воронки уже известен, его можно передать явно и пропустить запрос списка воронок:

```python

portal_360_deals = collect_deals_dataframe(category_id=7)
```
