import os
import pandas as pd
from functions import (
    prepare_directories,
    setup_logging,
    log,
    process_module,
    combine_processed_files,
    smart_merge,
    set_log_folder
)

# --- Константы и конфигурация ---
INPUT_FOLDER = r"C:\Users\geg\Desktop\Новая папка\Обрабатываемые"
PROCESSED_FOLDER = os.path.join(os.path.dirname(INPUT_FOLDER), "Обработанные")
LOG_FOLDER = os.path.join(os.path.dirname(INPUT_FOLDER), "log")
set_log_folder(LOG_FOLDER)

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

# --- Подготовка ---
prepare_directories([INPUT_FOLDER, PROCESSED_FOLDER, LOG_FOLDER])
setup_logging(LOG_FOLDER)

# --- Обработка каждого модуля ---
for key, config in MODULES.items():
    df = process_module(INPUT_FOLDER, PROCESSED_FOLDER,
                        key, config, RENAME_MAP)
    if df is not None:
        output_file = os.path.join(PROCESSED_FOLDER, f"Обработано_{key}.xlsx")
        df.to_excel(output_file, index=False)
        log(f"Результат сохранён: {output_file}")

# --- Объединение обработанных таблиц ---
processed_dfs = []
for key in MODULES.keys():
    file_path = os.path.join(PROCESSED_FOLDER, f"Обработано_{key}.xlsx")
    if os.path.exists(file_path):
        df = pd.read_excel(file_path)
        processed_dfs.append((key, df))

combined = combine_processed_files(processed_dfs)

# --- Сохранение до удаления дубликатов ---
pre_dedup_path = os.path.join(PROCESSED_FOLDER, "до_удаления_дубликатов.xlsx")
combined.to_excel(pre_dedup_path, index=False)
log(f"Промежуточный файл сохранён: {pre_dedup_path}")

# --- Удаление дубликатов ---
final_df = smart_merge(combined)

# --- Сохранение финального файла ---
final_path = os.path.join(PROCESSED_FOLDER, "Общий_итог.xlsx")
final_df.to_excel(final_path, index=False)
log(f"Финальный объединённый файл сохранён: {final_path}")
