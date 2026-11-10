"""Шаг 0. Проверка окружения.

Скрипт делает три вещи:
1. печатает версию Python и версии библиотек из requirements.txt;
2. собирает маленькую демонстрационную таблицу в pandas;
3. считает среднюю цену по брендам (как пример GROUP BY).

Аналогия для Java-разработчика:
- DataFrame — это `List<Row>` + именованные колонки, поверх которого работают
  SQL-подобные операции;
- ``groupby(...).agg(...)`` — это `Collectors.groupingBy` + `averagingDouble`;
- ``.round(2)`` применяется к каждой ячейке таблицы, как `map` по значениям.

Запуск:  python src/check_env.py
"""

from __future__ import annotations

import importlib
import platform
import sys

import pandas as pd

# Пакет для импорта -> имя пакета в pip (у некоторых они различаются).
REQUIRED_PACKAGES: dict[str, str] = {
    "pandas": "pandas",
    "numpy": "numpy",
    "scipy": "scipy",
    "statsmodels": "statsmodels",
    "sklearn": "scikit-learn",
    "matplotlib": "matplotlib",
    "seaborn": "seaborn",
    "pyarrow": "pyarrow",
    "docx": "python-docx",
    "pptx": "python-pptx",
}


def check_versions() -> None:
    """Печатает версии Python и библиотек, проверяя, что все они импортируются."""
    print("=" * 62)
    print(f"Python {platform.python_version()}  ({platform.system()} {platform.machine()})")
    print(f"Интерпретатор: {sys.executable}")
    print("=" * 62)

    missing: list[str] = []
    for module_name, pip_name in REQUIRED_PACKAGES.items():
        try:
            module = importlib.import_module(module_name)
        except ImportError:
            missing.append(pip_name)
            print(f"  [НЕТ]      {pip_name}")
            continue
        version = getattr(module, "__version__", "версия не указана")
        print(f"  [ok]       {pip_name:<15} {version}")

    if missing:
        # Явно падаем с ненулевым кодом, чтобы проблему было видно в CI и терминале.
        raise SystemExit(
            "Не хватает пакетов: "
            + ", ".join(missing)
            + "\nУстановите их: pip install -r requirements.txt"
        )
    print()


def build_demo_table() -> pd.DataFrame:
    """Возвращает маленькую таблицу «товар — бренд — цена».

    Настоящие датасеты появятся только на Шаге 1, поэтому здесь синтетика:
    она нужна исключительно для проверки, что pandas работает.
    """
    data = {
        "товар": [
            "Смартфон A1", "Смартфон A2", "Смартфон B1",
            "Ноутбук C1", "Ноутбук C2", "Ноутбук D1",
            "Наушники E1", "Наушники E2", "Наушники F1",
        ],
        "бренд": [
            "Alpha", "Alpha", "Beta",
            "Beta", "Gamma", "Gamma",
            "Alpha", "Beta", "Gamma",
        ],
        "категория": [
            "смартфоны", "смартфоны", "смартфоны",
            "ноутбуки", "ноутбуки", "ноутбуки",
            "наушники", "наушники", "наушники",
        ],
        "цена": [28990.0, 34990.0, 31990.0, 74990.0, 89990.0, 67990.0, 4990.0, 6990.0, 5490.0],
    }
    return pd.DataFrame(data)


def brand_price_stats(df: pd.DataFrame) -> pd.DataFrame:
    """Средняя, медианная цена и число товаров по каждому бренду."""
    stats = (
        df.groupby("бренд", as_index=False)
        .agg(
            товаров=("товар", "count"),
            средняя_цена=("цена", "mean"),
            медианная_цена=("цена", "median"),
        )
        .sort_values("средняя_цена", ascending=False)
        .round(2)
    )
    return stats


def main() -> None:
    check_versions()

    df = build_demo_table()
    print("Демонстрационная таблица:")
    print(df.to_string(index=False))
    print()

    print("Средние цены по брендам (аналог GROUP BY в SQL):")
    print(brand_price_stats(df).to_string(index=False))
    print()
    print("Окружение готово к работе.")


if __name__ == "__main__":
    main()
