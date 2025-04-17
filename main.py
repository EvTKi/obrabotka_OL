import os
from datetime import datetime
import pandas as pd
from functions import (
    load_named_table,
    combine_dataframes,
    save_dataframe_to_excel,
    smart_merge,
    apply_replacements,
    combine_columns_by_replace_key
)
from logger_utils import (set_log_file_path, set_log_level)
from config_manager import config
from pathlib import Path
import logging

# --- Константы ---
INPUT_FOLDER = config.ROOT / "Обрабатываемые"
PROCESSED_FOLDER = config.ROOT / "Обработанные"
LOG_FOLDER = config.ROOT / "log"

# Создаем папку для обработанных файлов
os.makedirs(PROCESSED_FOLDER, exist_ok=True)
os.makedirs(LOG_FOLDER, exist_ok=True)  # Создаем папку для логов

log_filename = f"log {datetime.now().strftime('%Y-%m-%d %H-%M-%S')}.log"
# Используем Path для формирования пути
log_file_path = LOG_FOLDER / log_filename
set_log_file_path(str(log_file_path))  # Конвертируем в строку перед передачей
set_log_level(2)
RENAME_MAP = config.RENAME_MAP
REPLACE_ENERGYMAIN = config.REPLACE_ENERGYMAIN
REPLACE_ACCESS = config.REPLACE_ACCESS
MODULES = config.MODULES
MERGE_REPLACEMENTS = {**REPLACE_ENERGYMAIN, **REPLACE_ACCESS}


# replace_energymain_keys = list(config.REPLACE_ENERGYMAIN.keys())
# replace_access_keys = list(config.REPLACE_ACCESS.keys())
# print(replace_energymain_keys)

# --- Основная обработка ---
all_dfs = []
all_combined_data = []  # Для сохранения исходных данных до smart_merge

for module_key, module_config in MODULES.items():
    table_names = module_config["table_names"]
    columns_to_remove = module_config["columns_to_remove"]

    dfs = []
    for filename in os.listdir(INPUT_FOLDER):
        if not filename.endswith(".xlsx"):
            continue

        if "~$" in filename or "log" in filename or "Обработано" in filename:
            continue

        # Если ключ модуля не совпадает, пропускаем файл
        if module_key not in filename:
            continue

        file_path = INPUT_FOLDER / filename  # Используем Path для формирования пути
        for table_name in table_names:
            df = load_named_table(file_path, table_name)
            if df is not None:
                dfs.append(df)

    if not dfs:
        print(
            f"Не загружено ни одной таблицы из файла {filename} для модуля {module_key}")
        logging.warning(
            f"Не загружено ни одной таблицы из файла {filename} для модуля {module_key}")
        continue

    # Объединяем все DataFrame
    combined_df = combine_dataframes(dfs, columns_to_remove, RENAME_MAP)

    # Добавляем объединённые данные в список для сохранения
    all_combined_data.append(combined_df)

    # Сохраняем промежуточный результат для каждого модуля
    processed_path = PROCESSED_FOLDER / f"Обработано_{module_key}.xlsx"
    save_dataframe_to_excel(combined_df, str(processed_path))

    all_dfs.append(combined_df)

if not all_dfs:
    print("Нет данных для объединения.")
    exit()

# Объединяем все DataFrame в один
final_combined_df = pd.concat(all_dfs, ignore_index=True)

# Применение замен для REPLACE_ENERGYMAIN
final_combined_df = apply_replacements(final_combined_df, REPLACE_ENERGYMAIN)

# Применение замен для REPLACE_ACCESS
final_combined_df = apply_replacements(final_combined_df, REPLACE_ACCESS)

# Сохраняем итоговый файл до применения smart_merge
final_path = PROCESSED_FOLDER / "итог_до_удаления_дубликатов.xlsx"
save_dataframe_to_excel(final_combined_df, str(final_path))

# Применение smart_merge для итогового DataFrame
final_combined_df = smart_merge(final_combined_df, RENAME_MAP)

# Сохраняем итоговый файл после применения smart_merge (удаления дубликатов)
final_path = PROCESSED_FOLDER / "итог_после_удаления_дубликатов.xlsx"
save_dataframe_to_excel(final_combined_df, str(final_path))
# Применяем combine_columns_by_replace_key для energymain
final_combined_df = combine_columns_by_replace_key(final_combined_df,
                                                   "REPLACE_ENERGYMAIN",
                                                   config,
                                                   drop=True)

# Сохраняем после объединения столбцов energymain
final_path = PROCESSED_FOLDER / \
    "итог_после_объединения_energymain.xlsx"
save_dataframe_to_excel(final_combined_df, str(final_path))
# Применяем combine_columns_by_replace_key для access
final_combined_df = combine_columns_by_replace_key(final_combined_df,
                                                   "REPLACE_ACCESS",
                                                   config)

# Сохраняем после объединения столбцов access
final_path = PROCESSED_FOLDER / \
    "итог_после_объединения_access.xlsx"

save_dataframe_to_excel(final_combined_df, str(final_path))
