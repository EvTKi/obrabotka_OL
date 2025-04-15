import os
from datetime import datetime
import pandas as pd
from functions import (
    load_named_table,
    combine_dataframes,
    save_dataframe_to_excel,
    smart_merge,
    apply_replacements
)
from logger_utils import (
    set_log_file_path
)
import json
from pathlib import Path

# --- Константы ---
ROOT = Path.cwd()
INPUT_FOLDER = ROOT / "Обрабатываемые"
PROCESSED_FOLDER = ROOT / "Обработанные"
LOG_FOLDER = ROOT / "log"

os.makedirs(PROCESSED_FOLDER, exist_ok=True)
os.makedirs(LOG_FOLDER, exist_ok=True)

log_filename = f"log {datetime.now().strftime('%Y-%m-%d %H-%M-%S')}.log"
log_file_path = os.path.join(LOG_FOLDER, log_filename)
set_log_file_path(log_file_path)

with open("config.json", "r", encoding="utf-8") as f:
    config = json.load(f)

RENAME_MAP = config["RENAME_MAP"]
REPLACE_ENERGYMAIN = config["REPLACE_ENERGYMAIN"]
MODULES = config["MODULES"]

# --- Основная обработка ---
all_dfs = []

for filename in os.listdir(INPUT_FOLDER):
    if not filename.endswith(".xlsx"):
        continue

    if "~$" in filename or "log" in filename or "Обработано" in filename:
        continue

    module_key = next((key for key in MODULES if key in filename), None)

    if not module_key:
        print(f"Пропущен файл без описанного модуля: {filename}")
        continue  # <-- Пропускаем файл

    file_path = os.path.join(INPUT_FOLDER, filename)
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

# Применение замен
final_df = apply_replacements(final_df, REPLACE_ENERGYMAIN)

final_path = os.path.join(PROCESSED_FOLDER, "итог.xlsx")
save_dataframe_to_excel(final_df, final_path)
