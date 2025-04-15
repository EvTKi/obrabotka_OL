# logger_utils.py

import logging

log_file_path = None


def set_log_file_path(path):
    global log_file_path
    log_file_path = path


def set_log_level(level: int):
    """
    Устанавливает уровень логирования:
    1 — только ошибки (ERROR),
    2 — стандартное логирование (INFO),
    3 — подробное логирование (DEBUG)
    """
    if level == 1:
        log_level = logging.ERROR
    elif level == 2:
        log_level = logging.INFO
    else:
        log_level = logging.DEBUG

    logging.basicConfig(
        filename=log_file_path,
        level=log_level,
        format="%(asctime)s - %(levelname)s - %(message)s"
    )


def log_decorator(func):
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except ValueError as ve:
            logging.warning(f"Предупреждение в функции {func.__name__}: {ve}")
            raise
        except Exception as e:
            logging.exception(f"Ошибка в функции {func.__name__}: {e}")
            raise
    return wrapper
