import os
import logging
import pandas as pd
from openpyxl import load_workbook

log_file_path = None


def set_log_file_path(path):
    global log_file_path
    log_file_path = path
    logging.basicConfig(
        filename=log_file_path,
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s"
    )


def log_decorator(func):
    def wrapper(*args, **kwargs):
        try:
            result = func(*args, **kwargs)
            return result
        except Exception as e:
            msg = f"Ошибка в функции {func.__name__}: {e}"
            print(msg)
            if log_file_path:
                logging.exception(msg)
            raise
    return wrapper


@log_decorator
def load_named_table(filepath, table_name):
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
                    print(
                        f"Загружена таблица '{table_name}' из файла {os.path.basename(filepath)}")
                    if log_file_path:
                        logging.info(
                            f"Загружена таблица '{table_name}' из файла {os.path.basename(filepath)}")
                    return df
        raise ValueError(
            f"Таблица с именем '{table_name}' не найдена в файле {filepath}")
    except Exception as e:
        raise RuntimeError(
            f"Ошибка при загрузке таблицы '{table_name}' из файла {filepath}: {e}")


@log_decorator
def load_all_tables_from_file(filepath):
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

        print(f"В файле '{filepath}' найдены таблицы: \n{all_table_names}")
        if log_file_path:
            logging.info(
                f"В файле '{filepath}' найдены таблицы: \n{all_table_names}")
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
def save_dataframe_to_excel(df, path):
    df.to_excel(path, index=False)
    print(f"Результат сохранён: {path}")
    if log_file_path:
        logging.info(f"Результат сохранён: {path}")


def merge_group(group):
    return group.apply(lambda col: next((x for x in col if pd.notna(x) and str(x).strip() != ''), None))


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
