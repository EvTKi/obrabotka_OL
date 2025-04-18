import customtkinter as ctk
from tkinter import filedialog, messagebox
import subprocess
import os
from config_manager import config
import json
import shutil


class AdminHelperApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Помощник Администратора")
        self.geometry("900x600")
        ctk.set_default_color_theme("blue")
        ctk.set_appearance_mode("light")

        self.input_folder = ctk.StringVar()

        self.build_sidebar()
        self.build_main_area()
        self.settings_window = None
        self.rename_entries = {}

    def build_sidebar(self):
        sidebar = ctk.CTkFrame(
            self, width=160, fg_color="#004C99", corner_radius=0)
        sidebar.pack(side="left", fill="y")

        settings_btn = ctk.CTkButton(
            sidebar,
            text="Настройки",
            fg_color="#3399FF",
            text_color="white",
            hover_color="#2A85D0",
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.open_settings
        )
        settings_btn.pack(pady=20, padx=10, fill="x")

    def build_main_area(self):
        main_frame = ctk.CTkFrame(self, fg_color="white")
        main_frame.pack(side="right", fill="both",
                        expand=True, padx=20, pady=20)
        main_frame.pack_propagate(False)

        label = ctk.CTkLabel(
            main_frame, text="Каталог с данными:", text_color="black")
        label.pack(anchor="nw")

        folder_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        folder_frame.pack(fill="x", pady=(10, 20))

        folder_entry = ctk.CTkEntry(
            folder_frame, textvariable=self.input_folder, state="readonly", width=500)
        folder_entry.pack(side="left", padx=(0, 10), fill="x", expand=True)

        browse_btn = ctk.CTkButton(
            folder_frame, text="Выбрать...", command=self.select_input_folder)
        browse_btn.pack(side="left")

        run_btn = ctk.CTkButton(
            main_frame,
            text="Запуск",
            fg_color="#007ACC",
            hover_color="#005F9E",
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.run_main_script
        )
        run_btn.place(relx=1.0, rely=1.0, anchor="se", x=-10, y=-10)

    def select_input_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.input_folder.set(folder)
            os.environ["INPUT_FOLDER"] = folder

    def run_main_script(self):
        if not self.input_folder.get():
            messagebox.showwarning(
                "Предупреждение", "Сначала выберите каталог с данными.")
            return
        try:
            subprocess.run(["python", "main.py"], check=True)
            messagebox.showinfo("Успех", "main.py успешно выполнен.")
        except subprocess.CalledProcessError as e:
            messagebox.showerror(
                "Ошибка", f"main.py завершился с ошибкой:\n{e}")

    def open_settings(self):
        if self.settings_window is not None and self.settings_window.winfo_exists():
            self.settings_window.focus()
            return

        self.settings_window = ctk.CTkToplevel(self)
        self.settings_window.title("Настройки")
        self.settings_window.geometry("700x500")
        self.settings_window.transient(self)
        self.settings_window.attributes("-topmost", True)
        self.settings_window.grab_set()
        self.settings_window.focus()
        self.settings_window.protocol(
            "WM_DELETE_WINDOW", self.on_close_settings)

        self.tabview = ctk.CTkTabview(
            self.settings_window, width=680, height=420)
        self.tabview.pack(padx=10, pady=10, fill="both", expand=True)

        self.rename_entries = {}

        self.config_data = config.read_raw()  # читаем полный config.json

        for section in self.config_data:
            self.build_config_tab(section, self.config_data[section])

        save_button = ctk.CTkButton(
            self.settings_window, text="Сохранить", command=self.save_config)
        save_button.pack(pady=10, padx=10, anchor="se")

    def build_config_tab(self, tab_name, mapping):
        tab = self.tabview.add(tab_name)

        header_frame = ctk.CTkFrame(tab)
        header_frame.pack(fill="x", pady=(10, 0))
        ctk.CTkLabel(header_frame, text="Ключ", width=250).pack(
            side="left", padx=(10, 0))
        ctk.CTkLabel(header_frame, text="Значение").pack(side="left", padx=10)

        self.rename_entries[tab_name] = {}

        for key, value in mapping.items():
            row = ctk.CTkFrame(tab)
            row.pack(fill="x", pady=2, padx=10)

            orig_entry = ctk.CTkEntry(row, width=250)
            orig_entry.insert(0, key)
            orig_entry.pack(side="left")

            val_entry = ctk.CTkEntry(row, width=300)
            val_entry.insert(0, str(value))
            val_entry.pack(side="left", padx=10)

            self.rename_entries[tab_name][orig_entry] = val_entry

    def save_config(self):
        updated_config = {}

        for tab_name, entries in self.rename_entries.items():
            updated_tab = {}
            for orig_entry, val_entry in entries.items():
                key = orig_entry.get()
                value = val_entry.get()
                if key:
                    updated_tab[key] = self.parse_value(value)
            updated_config[tab_name] = updated_tab

        with open("config_remake.json", "w", encoding="utf-8") as f:
            json.dump(updated_config, f, ensure_ascii=False, indent=4)

        ctk.CTkLabel(self.settings_window, text="Изменения сохранены!",
                     text_color="green").pack(pady=5)

    def parse_value(self, value):
        try:
            # Пробуем привести к числу или списку
            return json.loads(value)
        except:
            return value

    def on_close_settings(self):
        if self.settings_window is not None:
            self.settings_window.destroy()
            self.settings_window = None


if __name__ == "__main__":
    app = AdminHelperApp()
    app.mainloop()
