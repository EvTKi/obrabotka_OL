# functions.py

import os
import pandas as pd
from openpyxl import load_workbook
import logging
from logger_utils import log_decorator


@log_decorator
def load_named_table(filepath, table_name, verbose=False):
    try:
        wb = load_workbook(filepath, data_only=True)
        for sheet in wb.worksheets:
            for tbl in sheet._tables.values():
                if tbl.name == table_name:
                    data = sheet[tbl.ref]
                    rows = [[cell.value for cell in row] for row in data]
                    headers = rows[0]
                    values = rows[1:]
                    df = pd.DataFrame(values, columns=headers)
                    msg = f"Загружена таблица '{table_name}' из файла {os.path.basename(filepath)}"
                    if verbose:
                        print(msg)
                    logging.info(msg)
                    return df
        raise ValueError(
            f"Таблица с именем '{table_name}' не найдена в файле {filepath}")
    except Exception as e:
        raise RuntimeError(f"Ошибка при загрузке таблицы: {e}")


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
def smart_merge(df, rename_map):
    df = df.copy()
    df.rename(columns=rename_map, inplace=True)
    df.dropna(subset=['ФИО', 'УЗ'], how='all', inplace=True)

    def merge_group(group):
        return next((x for x in group if pd.notna(x) and str(x).strip() != ''), None)

    grouped = df.groupby(['ФИО', 'УЗ'], dropna=False).agg(
        merge_group).reset_index()
    return grouped


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
