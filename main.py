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

combined = combine_processed_files(processed_dfs)

# Сохранение объединённого файла ДО удаления дубликатов
pre_dedup_path = os.path.join(PROCESSED_FOLDER, "до_удаления_дубликатов.xlsx")
combined.to_excel(pre_dedup_path, index=False)
log(f"Промежуточный файл сохранён: {pre_dedup_path}")

# --- Шаг 4. Обработка дубликатов по 'УЗ' и 'ФИО' ---


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


# Финальное удаление дубликатов
final_df = smart_merge(combined)

# Сохранение объединённого финального файла
final_path = os.path.join(PROCESSED_FOLDER, "Общий_итог.xlsx")
final_df.to_excel(final_path, index=False)
log(f"Финальный объединённый файл сохранён: {final_path}")
