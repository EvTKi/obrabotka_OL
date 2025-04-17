# config_manager.py
from pathlib import Path
import json
from typing import Dict, Any


class AppConfig:
    def __init__(self):
        self.ROOT = Path.cwd()
        self.INPUT_FOLDER = self.ROOT / "Обрабатываемые"
        self.PROCESSED_FOLDER = self.ROOT / "Обработанные"
        self.LOG_FOLDER = self.ROOT / "log"
        self._load_config()

    def _load_config(self) -> None:
        with open("config.json", "r", encoding="utf-8") as f:
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


# Глобальный инстанс конфига
config = AppConfig()
