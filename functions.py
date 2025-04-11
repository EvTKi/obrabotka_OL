import os
import pandas as pd
import logging
from functools import wraps


def log(message):
    print(message)
    logging.info(message)


def trace(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        log(f"Вызов функции: {func.__name__}")
        result = func(*args, **kwargs)
        return result
    return wrapper


@trace
def prepare_directories(paths):
    for path in paths:
        os.makedirs(path, exist_ok=True)


@trace
def setup_logging(log_folder):
    log_file = os.path.join(log_folder, "pipeline.log")
    logging.basicConfig(
        filename=log_file,
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s"
    )


@trace
def process_module(input_folder, processed_folder, key, config):
    dfs = []
    for table_name in config["table_names"]:
        filename = f"Опросный лист {key}.xlsx"
        file_path = os.path.join(input_folder, filename)

        if not os.path.exists(file_path):
            log(f"Файл не найден и пропущен: {file_path}")
            continue

        try:
            df = pd.read_excel(file_path, sheet_name=table_name)
        except ValueError:
            log(f"Лист с именем '{table_name}' не найден в файле {file_path}")
            continue

        df["Модуль"] = key
        dfs.append(df)

    if not dfs:
        log(f"Нет данных для модуля: {key}")
        return None

    df = pd.concat(dfs, ignore_index=True)

    # Удаляем ненужные столбцы, если они есть
    columns_to_remove = config.get("columns_to_remove", [])
    df = df.drop(
        columns=[col for col in columns_to_remove if col in df.columns], errors="ignore")

    print(f"Столбцы после удаления: {df.columns}")
    print(f"Размер после объединения: {df.shape}")

    return df


@trace
def combine_processed_files(processed_dfs):
    all_data = []
    for key, df in processed_dfs:
        df["Модуль"] = key
        all_data.append(df)
    combined = pd.concat(all_data, ignore_index=True)
    print("Комбинированная таблица создана.")

    return combined


@trace
def smart_merge(df):
    key_column = "УЗ|ФИО"

    if key_column not in df.columns:
        fio = df.get("ФИО")
        account = df.get("Учетная запись")

        if fio is not None and account is not None:
            df[key_column] = df["Учетная запись"].astype(
                str) + "|" + df["ФИО"].astype(str)
        else:
            raise KeyError(
                "Нет столбца 'УЗ|ФИО' и невозможно создать его — отсутствуют 'ФИО' или 'Учетная запись'")

    df = df.sort_values(key_column)
    before = df.shape[0]
    df = df.drop_duplicates(subset=key_column, keep="first")
    after = df.shape[0]

    log(
        f"Удалено дубликатов по '{key_column}'. Итоговая форма: ({after}, {df.shape[1]})")
    return df
