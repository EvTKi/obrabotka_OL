# functions.py

import os
import pandas as pd
from openpyxl import load_workbook
import logging
from logger_utils import log_decorator
from functools import lru_cache


@lru_cache(maxsize=10)  # Кэширует 10 последних файлов
def _load_workbook_cached(filepath: str):
    """Кэшированная версия load_workbook"""
    return load_workbook(filepath, data_only=True)


@log_decorator
def load_named_table(filepath: str, table_name: str) -> pd.DataFrame:
    try:
        wb = _load_workbook_cached(filepath)  # Используем кэшированную версию
        for sheet in wb.worksheets:
            for tbl in sheet._tables.values():
                if tbl.name == table_name:
                    data = sheet[tbl.ref]
                    rows = [[cell.value for cell in row] for row in data]
                    return pd.DataFrame(rows[1:], columns=rows[0])
        raise ValueError(f"Таблица '{table_name}' не найдена")
    except Exception as e:
        raise RuntimeError(f"Ошибка загрузки: {e}")


@log_decorator
def load_all_tables_from_file(filepath, verbose=False):
    tables = []
    try:
        wb = load_workbook(filepath, data_only=True)
        all_table_names = []

        for sheet in wb.worksheets:
            for tbl in sheet._tables.values():
                all_table_names.append(tbl.name)
                data = sheet[tbl.ref]
                rows = [[cell.value for cell in row] for row in data]
                headers = rows[0]
                values = rows[1:]
                df = pd.DataFrame(values, columns=headers)
                tables.append(df)

        msg = f"В файле '{filepath}' найдены таблицы: \n{all_table_names}"
        if verbose:
            print(msg)
        logging.info(msg)
        return tables
    except Exception as e:
        raise RuntimeError(
            f"Ошибка при чтении всех таблиц из файла {filepath}: {e}")


@log_decorator
def combine_dataframes(dfs, columns_to_remove, rename_map):
    combined = pd.concat(dfs, ignore_index=True)
    combined = combined.copy()
    combined.drop(columns=[
                  col for col in columns_to_remove if col in combined.columns], errors='ignore', inplace=True)
    combined.rename(columns=rename_map, inplace=True)
    return combined


@log_decorator
def save_dataframe_to_excel(df, path, verbose=False):
    df.to_excel(path, index=False)
    msg = f"Результат сохранён: {path}"
    if verbose:
        print(msg)
    logging.info(msg)


@log_decorator
def smart_merge(df: pd.DataFrame, rename_map: dict[str, str]) -> pd.DataFrame:
    """
    Удаляет дубликаты по комбинации столбцов 'ФИО' и 'УЗ', сохраняя первое вхождение.

    Args:
        df: Входной DataFrame
        rename_map: Словарь для переименования столбцов

    Returns:
        DataFrame без дубликатов
    """
    df = df.copy()
    df.rename(columns=rename_map, inplace=True)

    # Оптимизация через drop_duplicates
    if all(col in df.columns for col in ['ФИО', 'УЗ']):
        return df.drop_duplicates(subset=['ФИО', 'УЗ'], keep='first')
    return df


@log_decorator
def apply_replacements(df: pd.DataFrame, replace_dict: dict, verbose=False) -> pd.DataFrame:
    """
    Применяет замены значений в указанных столбцах DataFrame согласно словарю replacements.
    Формат словаря:
    {
        "Имя столбца": {
            "старое значение": "новое значение",
            ...
        },
        ...
    }
    """
    df = df.copy()
    for column, replacements in replace_dict.items():
        if column in df.columns:
            for old_value, new_value in replacements.items():
                mask = df[column] == old_value
                if mask.any():
                    msg = f"Значение столбца '{column}' заменено с '{old_value}' на '{new_value}'"
                    if verbose:
                        print(msg)
                    logging.info(msg)
                    df.loc[mask, column] = new_value
        else:
            msg = f"Столбец '{column}' не найден в DataFrame для замены."
            if verbose:
                print(msg)
            logging.warning(msg)
    return df
