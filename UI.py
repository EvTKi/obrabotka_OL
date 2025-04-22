import customtkinter as ctk
from tkinter import filedialog, messagebox
import tkinter as tk
import subprocess
import os
import json
from config_manager import config


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
        self.settings_window.geometry("750x520")
        self.settings_window.transient(self)
        self.settings_window.attributes("-topmost", True)
        self.settings_window.grab_set()
        self.settings_window.focus()
        self.settings_window.protocol(
            "WM_DELETE_WINDOW", self.on_close_settings)

        self.tabview = ctk.CTkTabview(
            self.settings_window, width=700, height=420)
        self.tabview.pack(padx=10, pady=(10, 0), fill="both", expand=True)

        self.rename_entries = {}

        self.config_data = config.read_raw()

        for section in self.config_data:
            self.build_config_tab(section, self.config_data[section])

        # Кнопки в одну строку снизу
        button_frame = ctk.CTkFrame(self.settings_window)
        button_frame.pack(pady=10, padx=10, anchor="se", fill="x")

        self.success_label = ctk.CTkLabel(
            button_frame, text="", text_color="green")
        self.success_label.pack(side="left", padx=(10, 0))

        add_btn = ctk.CTkButton(
            button_frame, text="Добавить строку", command=self.add_entry_to_active_tab)
        add_btn.pack(side="right", padx=5)

        del_btn = ctk.CTkButton(button_frame, text="Удалить строку",
                                command=self.remove_last_entry_from_active_tab)
        del_btn.pack(side="right", padx=5)

        save_btn = ctk.CTkButton(
            button_frame, text="Сохранить", command=self.save_config)
        save_btn.pack(side="right", padx=5)

    def build_config_tab(self, tab_name, mapping):
        tab = self.tabview.add(tab_name)

        canvas = tk.Canvas(tab)
        scrollbar = ctk.CTkScrollbar(tab, command=canvas.yview)
        scrollable_frame = ctk.CTkFrame(canvas)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # ✅ Привязка прокрутки к canvas, а не глобально
        canvas.bind("<MouseWheel>", lambda event: canvas.yview_scroll(
            int(-1 * (event.delta / 120)), "units"))

        header_frame = ctk.CTkFrame(scrollable_frame)
        header_frame.pack(fill="x", pady=(10, 0))
        ctk.CTkLabel(header_frame, text="Ключ", width=250).pack(
            side="left", padx=(10, 0))
        ctk.CTkLabel(header_frame, text="Значение").pack(side="left", padx=10)

        self.rename_entries[tab_name] = {
            "frame": scrollable_frame, "entries": []}

        for key, value in mapping.items():
            self.add_entry_to_tab(tab_name, key, str(value))

    def add_entry_to_tab(self, tab_name, key="", value=""):
        frame = self.rename_entries[tab_name]["frame"]
        row = ctk.CTkFrame(frame)
        row.pack(fill="x", pady=2, padx=10)

        orig_entry = ctk.CTkEntry(row, width=250)
        orig_entry.insert(0, key)
        orig_entry.pack(side="left")

        val_entry = ctk.CTkEntry(row, width=300)
        val_entry.insert(0, str(value))
        val_entry.pack(side="left", padx=10)

        self.rename_entries[tab_name]["entries"].append(
            (orig_entry, val_entry))
        row.update_idletasks()
        frame.update_idletasks()
        frame.master.yview_moveto(1.0)

    def add_entry_to_active_tab(self):
        current_tab = self.tabview.get()
        self.add_entry_to_tab(current_tab)

    def remove_last_entry_from_active_tab(self):
        current_tab = self.tabview.get()
        entries = self.rename_entries[current_tab]["entries"]
        if entries:
            entry_pair = entries.pop()
            entry_pair[0].master.destroy()

    def save_config(self):
        updated_config = {}

        for tab_name, data in self.rename_entries.items():
            updated_tab = {}
            for orig_entry, val_entry in data["entries"]:
                key = orig_entry.get()
                value = val_entry.get()
                if key:
                    updated_tab[key] = self.parse_value(value)
            updated_config[tab_name] = updated_tab

        with open("config.json", "w", encoding="utf-8") as f:
            json.dump(updated_config, f, ensure_ascii=False, indent=4)

        self.success_label.configure(text="Изменения сохранены!")

    def parse_value(self, value):
        try:
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
