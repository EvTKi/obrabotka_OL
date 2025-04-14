import os
import pandas as pd
import openpyxl
import datetime

# --- Глобальная переменная пути к лог-файлу ---
_log_file_path = None


def set_log_folder(log_folder):
    """
    Устанавливает путь к лог-файлу, создавая его в указанной папке с текущей датой и временем.
    """
    global _log_file_path
    timestamp = datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    _log_file_path = os.path.join(log_folder, f"log {timestamp}.log")


def setup_logging(log_folder):
    """
    Создаёт папку логов, если она не существует.
    """
    os.makedirs(log_folder, exist_ok=True)
    log("Логирование инициализировано.")


def log(message):
    """
    Записывает сообщение в лог-файл.
    """
    if _log_file_path is None:
        raise RuntimeError(
            "Лог-файл не инициализирован. Вызовите set_log_folder().")
    timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    full_message = f"[{timestamp}] {message}"
    with open(_log_file_path, "a", encoding="utf-8") as f:
        f.write(full_message + "\n")
    print(full_message)


def prepare_directories(folders):
    """
    Создаёт необходимые папки.
    """
    for folder in folders:
        os.makedirs(folder, exist_ok=True)
    log(f"Созданы папки: {folders}")


def load_named_table(file_path, table_name):
    """
    Загружает таблицу Excel с указанным именем из любого листа книги.
    Возвращает DataFrame, если таблица найдена. Иначе — вызывает ошибку.
    """
    wb = openpyxl.load_workbook(file_path, data_only=True)

    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        if table_name in sheet.tables:
            table = sheet.tables[table_name]
            data = sheet[table.ref]

            rows = [[cell.value for cell in row] for row in data]
            df = pd.DataFrame(rows[1:], columns=rows[0])
            return df

    raise KeyError(
        f"Таблица с именем '{table_name}' не найдена ни на одном листе в файле {file_path}.")


def process_module(input_folder, output_folder, key, config, rename_map):
    """
    Загружает и объединяет таблицы одного модуля, удаляет столбцы, переименовывает поля.
    """
    combined = []
    for filename in os.listdir(input_folder):
        if not filename.endswith(".xlsx") or key not in filename:
            continue

        file_path = os.path.join(input_folder, filename)
        for table in config["table_names"]:
            try:
                df = load_named_table(file_path, table)
                df.drop(columns=config["columns_to_remove"],
                        errors='ignore', inplace=True)
                df.rename(columns=rename_map, inplace=True)
                combined.append(df)
                log(f"Загружена таблица '{table}' из файла {filename}")
            except Exception as e:
                log(
                    f"Ошибка при загрузке таблицы '{table}' из файла {filename}: {e}")

    if not combined:
        log(f"Нет данных для модуля {key}")
        return None

    result = pd.concat(combined, ignore_index=True)
    return result


def combine_processed_files(dfs):
    """
    Объединяет предварительно обработанные таблицы из разных модулей.
    """
    combined = []
    for key, df in dfs:
        df["Модуль"] = key
        combined.append(df)
        log(f"Добавлен модуль '{key}' с {len(df)} строками")
    return pd.concat(combined, ignore_index=True)


def smart_merge(df):
    """
    Удаляет дубликаты по 'ФИО' и 'УЗ', исключает строки, где оба этих поля пустые.
    """
    df = df.copy()
    df = df[~((df["ФИО"].isna()) & (df["УЗ"].isna()))]
    df.drop_duplicates(subset=["ФИО", "УЗ"], inplace=True)
    log(f"После удаления дубликатов осталось строк: {len(df)}")
    return df
