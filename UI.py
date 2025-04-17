import customtkinter as ctk
from tkinter import filedialog, messagebox
import subprocess
import os


class AdminHelperApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Помощник Администратора")
        self.geometry("900x600")
        # можно также создать собственную тему
        ctk.set_default_color_theme("blue")
        ctk.set_appearance_mode("light")

        self.input_folder = ctk.StringVar()

        self.build_sidebar()
        self.build_main_area()

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
        # Позволяет фиксировать размер и позиционировать вручную
        main_frame.pack_propagate(False)

        # Поле выбора каталога
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

        # Кнопка запуска — вручную позиционируем в правом нижнем углу
        run_btn = ctk.CTkButton(
            main_frame,
            text="Запуск",
            fg_color="#007ACC",
            hover_color="#005F9E",
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.run_main_script
        )
        run_btn.place(relx=1.0, rely=1.0, anchor="se", x=-10,
                      y=-10)  # смещение от нижнего правого угла

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
        messagebox.showinfo("Настройки", "Окно настроек пока не реализовано.")


if __name__ == "__main__":
    app = AdminHelperApp()
    app.mainloop()
