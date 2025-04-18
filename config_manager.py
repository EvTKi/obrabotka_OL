# config_manager.py
from pathlib import Path
import json
from typing import Dict, Any
import shutil  # добавь к остальным импортам


class AppConfig:
    def __init__(self):
        self.ROOT = Path.cwd()
        self.INPUT_FOLDER = self.ROOT / "Обрабатываемые"
        self.PROCESSED_FOLDER = self.ROOT / "Обработанные"
        self.LOG_FOLDER = self.ROOT / "log"
        self._load_config()

    def _load_config(self) -> None:
        # Создаем config_remake.json, если его нет
        remake_path = self.ROOT / "config_remake.json"
        original_path = self.ROOT / "config.json"

        if not remake_path.exists():
            shutil.copy(original_path, remake_path)

        with open(original_path, "r", encoding="utf-8") as f:
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
        Возвращает полный JSON-конфиг (из config_remake.json, если он есть, иначе config.json).
        """
        path = self.ROOT / "config_remake.json"
        if not path.exists():
            path = self.ROOT / "config.json"

        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)


# Глобальный инстанс конфига
config = AppConfig()
