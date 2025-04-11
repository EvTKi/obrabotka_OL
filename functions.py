import os
import logging
import pandas as pd
from openpyxl import load_workbook
import numpy as np
from functools import wraps
from datetime import datetime


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


LOG_FOLDER = None  # будет задан из main


def set_log_folder(folder_path):
    global LOG_FOLDER
    LOG_FOLDER = folder_path


def log_decorator(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        if LOG_FOLDER is None:
            raise ValueError(
                "LOG_FOLDER не установлен. Используйте set_log_folder(path).")

        timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        log_filename = f"{func.__name__}_{timestamp}.log"
        log_path = os.path.join(LOG_FOLDER, log_filename)

        os.makedirs(LOG_FOLDER, exist_ok=True)
        with open(log_path, 'w', encoding='utf-8') as f:
            f.write(
                f"Вызов функции: {func.__name__} с аргументами: {args} {kwargs}\n")
            result = func(*args, **kwargs)
            f.write(f"Результат: {result}\n")
        return result
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
def smart_merge(df: pd.DataFrame) -> pd.DataFrame:
    """
    Объединяет строки по комбинации 'УЗ' и 'ФИО' с сохранением первых непустых значений.

    Этапы:
    1. Удаление строк, где оба поля 'ФИО' и 'УЗ' пустые.
    2. Создание уникального ключа на основе 'УЗ|ФИО'.
    3. Группировка по ключу и объединение строк с приоритетом непустых значений.
    4. Очистка служебного ключа.

    :param df: DataFrame с колонками 'ФИО' и 'УЗ'
    :return: Обработанный DataFrame без дубликатов
    """
    # Шаг 1: удаляем строки, где оба поля пустые
    df = df.copy()
    mask = ~((df["ФИО"].fillna("") == "") & (df["УЗ"].fillna("") == ""))
    df = df[mask]

    # Шаг 2: создаём ключ для объединения
    df["_key"] = df["УЗ"].astype(str) + "|" + df["ФИО"].astype(str)

    print("Удаление дубликатов с переносом значений")

    # Шаг 3: группируем и сохраняем первые непустые значения
    df = df.groupby("_key").agg(lambda x: x.dropna(
    ).iloc[0] if x.dropna().any() else np.nan).reset_index(drop=True)

    # Шаг 4: удаляем ключ
    if "_key" in df.columns:
        df.drop(columns=["_key"], inplace=True)

    print(f"Удалено дубликатов по 'УЗ|ФИО'. Итоговая форма: {df.shape}")
    return df
