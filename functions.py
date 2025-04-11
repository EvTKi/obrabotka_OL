import os
import pandas as pd
from datetime import datetime
import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table

LOG_FILE = None


def prepare_directories(dirs):
    for directory in dirs:
        if not os.path.exists(directory):
            os.makedirs(directory)


def setup_logging(log_dir):
    global LOG_FILE
    now = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    LOG_FILE = os.path.join(log_dir, f"log_{now}.txt")
    with open(LOG_FILE, "w", encoding="utf-8") as f:
        f.write(f"Лог запущен: {now}\n")
    return LOG_FILE


def log(message):
    print(message)
    if LOG_FILE:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(message + "\n")


def process_module(input_folder, processed_folder, key, config):
    file_name = f"Опросный лист {key}.xlsx"
    file_path = os.path.join(input_folder, file_name)

    if not os.path.exists(file_path):
        log(f"Файл не найден и пропущен: {file_path}")
        return None

    dfs = []

    try:
        wb = load_workbook(file_path, data_only=True)

        # Перебираем таблицы по заданным именам
        for table_name in config["table_names"]:
            found = False
            for sheet in wb.worksheets:
                for tbl in sheet._tables.values() if isinstance(sheet._tables, dict) else sheet._tables:
                    tbl_obj = tbl if isinstance(
                        tbl, Table) else sheet.tables[tbl]

                    if tbl_obj.name == table_name:
                        # Получение диапазона таблицы
                        data_range = sheet[tbl_obj.ref]
                        headers = [cell.value for cell in data_range[0]]
                        rows = [[cell.value for cell in row]
                                for row in data_range[1:]]

                        df = pd.DataFrame(rows, columns=headers)

                        # Удаление ненужных столбцов
                        for col in config["columns_to_remove"]:
                            if col in df.columns:
                                df.drop(columns=[col], inplace=True)

                        df["Модуль"] = key
                        dfs.append(df)
                        found = True
                        break

                if found:
                    break

            if not found:
                log(f"Таблица '{table_name}' не найдена в файле: {file_path}")

    except Exception as e:
        log(f"Ошибка при обработке файла {file_path}: {e}")
        return None

    if dfs:
        result_df = pd.concat(dfs, ignore_index=True)
        return result_df
    else:
        log(f"Нет данных для модуля: {key}")
        return None


def combine_processed_files(processed_data):
    dfs = [df for _, df in processed_data]
    if dfs:
        return pd.concat(dfs, ignore_index=True)
    return pd.DataFrame()


def smart_merge(df):
    log("Удаление дубликатов с переносом значений")

    if "УЗ" not in df.columns or "ФИО" not in df.columns:
        log("Не найдены ключевые столбцы 'УЗ' или 'ФИО', пропуск удаления дубликатов")
        return df

    key_columns = ["УЗ", "ФИО"]
    grouped = df.groupby(key_columns, as_index=False)

    merged_rows = []

    for _, group in grouped:
        base = group.iloc[0].copy()

        for _, row in group.iloc[1:].iterrows():
            for col in df.columns:
                if col in key_columns:
                    continue
                base_val = base[col]
                new_val = row[col]

                if pd.isna(base_val) or base_val == 0 or base_val == "":
                    if not pd.isna(new_val) and new_val != 0 and new_val != "":
                        base[col] = new_val

        merged_rows.append(base)

    result_df = pd.DataFrame(merged_rows)
    log(f"Дубликаты объединены. Итоговая форма: {result_df.shape}")
    return result_df
