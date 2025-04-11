import os
from functions import (
    prepare_directories,
    setup_logging,
    process_module,
    combine_processed_files,
    smart_merge,
    log
)
import pandas as pd

# Конфигурация
# Укажи путь к папке с исходными данными
INPUT_FOLDER = r"C:\Users\geg\Desktop\Новая папка\Обрабатываемые"

# Эти папки автоматически создаются рядом с INPUT_FOLDER
PROCESSED_FOLDER = os.path.join(os.path.dirname(INPUT_FOLDER), "Обработанные")
LOG_FOLDER = os.path.join(os.path.dirname(INPUT_FOLDER), "log")

MODULES = {
    "ОЖ": {
        "table_names": ["ДП", "Рук", "Проч_персон"],
        "columns_to_remove": ["Столбец1"]
    },
    "ЖД": {
        "table_names": ["ЖД"],
        "columns_to_remove": ["Столбец1"]
    }
}

# Подготовка директорий
prepare_directories([INPUT_FOLDER, PROCESSED_FOLDER, LOG_FOLDER])
setup_logging(LOG_FOLDER)

# Обработка модулей
for key, config in MODULES.items():
    df = process_module(INPUT_FOLDER, PROCESSED_FOLDER, key, config)
    if df is not None:
        output_file = os.path.join(PROCESSED_FOLDER, f"Обработано_{key}.xlsx")
        df.to_excel(output_file, index=False)
        log(f"Результат сохранён: {output_file}")

# Объединение всех обработанных таблиц
processed_dfs = []
for key in MODULES.keys():
    file_path = os.path.join(PROCESSED_FOLDER, f"Обработано_{key}.xlsx")
    if os.path.exists(file_path):
        df = pd.read_excel(file_path)
        processed_dfs.append((key, df))

        print(f"--- Модуль {key} ---")
        print("Столбцы после process_module:", df.columns.tolist())
        print("Размер после process_module:", df.shape)
    else:
        print(f"Пропущен модуль {key}")

combined = combine_processed_files(processed_dfs)
print("Столбцы в combined:", combined.columns)
print("Размер combined:", combined.shape)

# Финальное удаление дубликатов
final_df = smart_merge(combined)

# Сохранение объединённого финального файла
final_path = os.path.join(PROCESSED_FOLDER, "Общий_итог.xlsx")
final_df.to_excel(final_path, index=False)
log(f"Финальный объединённый файл сохранён: {final_path}")
