import json
from pathlib import Path
from typing import Any


class AppConfig:
    """
    Класс для обработки конфигурационных данных из config.json.
    """

    # Предполагается, что config.json находится в корне проекта
    ROOT = Path(__file__).parent

    def __init__(self):
        path = self.ROOT / "config.json"
        with open(path, "r", encoding="utf-8") as f:
            config = json.load(f)

        self.RENAME_MAP: dict[str, str] = config["RENAME_MAP"]
        self.REPLACE_ENERGYMAIN: dict[str, Any] = config["REPLACE_ENERGYMAIN"]
        self.REPLACE_ACCESS: dict[str, Any] = config["REPLACE_ACCESS"]
        self.MODULES: dict[str, dict[str, Any]] = config["MODULES"]

    def get_config(self, key: str, default: Any = None):
        """
        Возвращает значение конфигурации по ключу.

        Args:
            key (str): Ключ для получения значения.
            default (Any): Значение по умолчанию, если ключ не найден.

        Returns:
            Значение конфигурации или значение по умолчанию.
        """
        # Проверка, есть ли атрибут с таким именем
        if hasattr(self, key):
            return getattr(self, key)
        return default

    def read_raw(self) -> dict:
        """
        Возвращает полный JSON-конфиг (из config.json).
        """
        path = self.ROOT / "config.json"

        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)

    def update_config(self, key: str, value: Any):
        """
        Обновляет значение конфигурации по ключу и сохраняет изменения в config.json.

        Args:
            key (str): Ключ для обновления значения.
            value (Any): Новое значение для сохранения.
        """
        path = self.ROOT / "config.json"

        with open(path, "r", encoding="utf-8") as f:
            config = json.load(f)

        config[key] = value

        with open(path, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4, ensure_ascii=False)

        # Обновляем атрибуты объекта
        setattr(self, key, value)


# Глобальный инстанс конфига
config = AppConfig()
