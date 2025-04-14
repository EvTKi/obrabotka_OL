import os
from datetime import datetime
import pandas as pd
from functions import (
    set_log_file_path,
    load_named_table,
    load_all_tables_from_file,
    combine_dataframes,
    save_dataframe_to_excel,
    smart_merge,
    set_log_level
)

# --- Константы ---
INPUT_FOLDER = os.path.join(os.getcwd(), "Обрабатываемые")
PROCESSED_FOLDER = os.path.join(os.getcwd(), "Обработанные")
LOG_FOLDER = os.path.join(os.getcwd(), "log")

os.makedirs(PROCESSED_FOLDER, exist_ok=True)
os.makedirs(LOG_FOLDER, exist_ok=True)

log_filename = f"log {datetime.now().strftime('%Y-%m-%d %H-%M-%S')}.log"
log_file_path = os.path.join(LOG_FOLDER, log_filename)
set_log_file_path(log_file_path)

LOGGING_LEVEL = 2  # 1 — ошибки, 2 — INFO, 3 — DEBUG
set_log_level(LOGGING_LEVEL)

# --- Словарь переименований ---
RENAME_MAP = {
    "ФИО сотрудника": "ФИО",
    "Сотрудник": "ФИО",
    "Пользователь": "Учетная запись",
    "Учетная запись сотрудника": "УЗ",
    "Учетная запись в MS AD": "УЗ",
    "Учетная запись MS AD": "УЗ",
    "Учетная запись в службе каталогов": "УЗ",
    "Электронная почта*": "Электронная почта",
    "Мобильный телефон*": "Мобильный телефон"
}

# --- Модули ---
MODULES = {
    "ОЖ": {
        "table_names": ["ДП", "Рук", "Проч_персон"],
        "columns_to_remove": ["Столбец1"]
    },
    "ЖД": {
        "table_names": ["ЖД"],
        "columns_to_remove": ["Столбец1"]
    },
    "ЖТАР": {
        "table_names": ["ГИД"],
        "columns_to_remove": ["Столбец1"]
    }
}

# --- Основная обработка --
all_dfs = []

for filename in os.listdir(INPUT_FOLDER):
    if not filename.endswith(".xlsx"):
        continue

    file_path = os.path.join(INPUT_FOLDER, filename)
    module_key = next((key for key in MODULES if key in filename), None)

    if not module_key:
        print(f"Модуль не найден для файла: {filename}")
        continue

    table_names = MODULES[module_key]["table_names"]
    columns_to_remove = MODULES[module_key]["columns_to_remove"]

    dfs = []

    for table_name in table_names:
        df = load_named_table(file_path, table_name)
        if df is not None:
            dfs.append(df)

    if not dfs:
        print(f"Не загружено ни одной таблицы из файла {filename}")
        continue

    combined_df = combine_dataframes(dfs, columns_to_remove, RENAME_MAP)
    processed_path = os.path.join(
        PROCESSED_FOLDER, f"Обработано_{module_key}.xlsx")
    save_dataframe_to_excel(combined_df, processed_path)

    all_dfs.append(combined_df)

if not all_dfs:
    print("Нет данных для объединения.")
    exit()

combined = pd.concat(all_dfs, ignore_index=True)
intermediate_path = os.path.join(
    PROCESSED_FOLDER, "до_удаления_дубликатов.xlsx")
save_dataframe_to_excel(combined, intermediate_path)

final_df = smart_merge(combined, RENAME_MAP)
final_path = os.path.join(PROCESSED_FOLDER, "итог.xlsx")
save_dataframe_to_excel(final_df, final_path)
