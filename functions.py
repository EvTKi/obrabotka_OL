import os
import logging
import pandas as pd
from openpyxl import load_workbook

logger = None


def setup_logging(log_folder):
    """
    Настраивает логирование в указанный файл внутри папки log_folder.
    """
    global logger
    os.makedirs(log_folder, exist_ok=True)
    log_file = os.path.join(log_folder, "log.txt")
    logging.basicConfig(filename=log_file, level=logging.INFO,
                        format="%(asctime)s - %(message)s")
    logger = logging.getLogger()
    return log_file


def log(message):
    """
    Печатает и логирует сообщение.
    """
    print(message)
    if logger:
        logger.info(message)


def log_decorator(func):
    """
    Декоратор для автоматического логирования вызова функции и её аргументов.
    """
    def wrapper(*args, **kwargs):
        log(f"Вызов функции: {func.__name__} с аргументами: {args} {kwargs}")
        return func(*args, **kwargs)
    return wrapper


@log_decorator
def prepare_directories(folders):
    """
    Создаёт указанные директории, если они не существуют.
    """
    for folder in folders:
        os.makedirs(folder, exist_ok=True)


@log_decorator
def load_named_table(file_path, table_name):
    """
    Загружает таблицу с указанным именем из Excel-файла.
    """
    wb = load_workbook(filename=file_path, data_only=True)
    for sheet in wb.worksheets:
        for tbl in sheet._tables.values():
            if tbl.name == table_name:
                data = sheet[tbl.ref]
                headers = [cell.value for cell in data[0]]
                rows = [[cell.value for cell in row] for row in data[1:]]
                df = pd.DataFrame(rows, columns=headers)
                return df
    raise ValueError(
        f"Таблица с именем '{table_name}' не найдена в файле '{file_path}'.")


@log_decorator
def rename_standard_columns(df, table_name, rename_map):
    """
    Переименовывает стандартные столбцы в DataFrame по предоставленной карте.
    """
    log(
        f"Переименование столбцов по стандартной карте для таблицы {table_name}: {rename_map}")
    return df.rename(columns=rename_map)


@log_decorator
def process_module(input_folder, processed_folder, key, config, rename_map):
    """
    Обрабатывает модуль: загружает указанные таблицы, переименовывает столбцы,
    удаляет лишние и добавляет информацию о модуле.
    """
    file_name = f"Опросный лист {key}.xlsx"
    file_path = os.path.join(input_folder, file_name)

    if not os.path.exists(file_path):
        log(f"Файл не найден и пропущен: {file_path}")
        return None

    df_list = []
    for table_name in config["table_names"]:
        try:
            df = load_named_table(file_path, table_name)
            df = rename_standard_columns(df, table_name, rename_map)

            if config.get("columns_to_remove"):
                df = df.drop(columns=[
                    col for col in config["columns_to_remove"] if col in df.columns
                ], errors='ignore')

            df["Модуль"] = key
            df_list.append(df)

        except Exception as e:
            log(f"Ошибка при чтении таблицы '{table_name}' в файле '{file_path}': {e}")

    if df_list:
        return pd.concat(df_list, ignore_index=True)
    else:
        log(f"Нет данных для модуля: {key}")
        return None


@log_decorator
def combine_processed_files(dataframes):
    """
    Объединяет список DataFrame в один.
    """
    combined = pd.concat([df for _, df in dataframes], ignore_index=True)
    return combined


@log_decorator
def smart_merge(df):
    """
    Удаляет дубликаты по 'УЗ' и 'ФИО', переносит значения из дубликатов в первую строку.
    Также удаляет строки, у которых 'УЗ' и 'ФИО' пустые.
    """
    log("Удаление дубликатов с переносом значений")

    if not {'УЗ', 'ФИО'}.issubset(df.columns):
        log("Не найдены ключевые столбцы 'УЗ' или 'ФИО', пропуск удаления дубликатов")
        return df

    df = df.copy()
    # удаляем строки с пустыми 'УЗ' и 'ФИО'
    df = df[~(df["УЗ"].isna() & df["ФИО"].isna())]
    df["_key"] = df["УЗ"].astype(str) + "|" + df["ФИО"].astype(str)
    grouped = df.groupby("_key", sort=False)

    merged_rows = []
    for _, group in grouped:
        base = group.iloc[0].copy()
        for _, row in group.iloc[1:].iterrows():
            for col in df.columns:
                if col not in ["_key", "УЗ", "ФИО"] and (
                    pd.isna(base[col]) or base[col] == 0 or base[col] == ""
                ):
                    base[col] = row[col]
        merged_rows.append(base)

    result_df = pd.DataFrame(merged_rows).drop(columns=["_key"])
    log(f"Удалено дубликатов по 'УЗ|ФИО'. Итоговая форма: {result_df.shape}")
    return result_df
