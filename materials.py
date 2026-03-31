import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import json
import os
from datetime import datetime, timedelta

class MaterialApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Учёт материалов — Единая база")
        self.root.geometry("1920x980")
        self.root.minsize(1400, 700)
        self.root.state('zoomed')  # Запуск в полноэкранном режиме по умолчанию

        self.data = []
        self.next_id = 0
        self.today = datetime.now().date()
        self.data_file = self.get_data_file()

        self.load_data()

        # ==================== МЕНЮ СВЕРХУ ====================
        menubar = tk.Menu(root)
        root.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Файл", menu=file_menu)
        file_menu.add_command(label="Сменить файл базы", command=self.change_data_file)
        file_menu.add_command(label="Экспорт в Excel", command=self.export_to_excel)
        file_menu.add_separator()
        file_menu.add_command(label="Выход", command=self.on_close)

        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Правка", menu=edit_menu)
        edit_menu.add_command(label="Добавить новую строку", command=self.add_new_row)
        edit_menu.add_command(label="Добавить новое поле", command=self.add_new_column)

        # ==================== ПАНЕЛЬ ИНСТРУМЕНТОВ ====================
        toolbar = tk.Frame(root, relief="raised", bd=1)
        toolbar.pack(fill="x", padx=5, pady=5)

        tk.Button(toolbar, text="➕ Добавить материал", width=20, height=2, bg="#4CAF50", fg="white",
                  font=("Arial", 10, "bold"), command=self.open_add_material_window).pack(side="left", padx=5)
        
        tk.Button(toolbar, text="📊 Экспорт в Excel", width=18, height=2, bg="#2196F3", fg="white",
                  command=self.export_to_excel).pack(side="left", padx=5)
        
        tk.Button(toolbar, text="🔄 Сменить базу", width=18, height=2, bg="#9C27B0", fg="white",
                  command=self.change_data_file).pack(side="left", padx=5)

        tk.Label(toolbar, text="   ").pack(side="left")  # отступ

        self.file_label = tk.Label(toolbar, text=f"Файл: {os.path.basename(self.data_file)}", 
                                  font=("Arial", 10, "bold"), fg="blue")
        self.file_label.pack(side="right", padx=15)

        # ==================== ТАБЛИЦА ====================
        self.tree = ttk.Treeview(root, show="headings", height=25)
        self.refresh_columns()

        scrollbar = ttk.Scrollbar(root, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side="left", fill="both", expand=True, padx=10, pady=5)
        scrollbar.pack(side="right", fill="y", padx=(0,10), pady=5)

        # Привязки
        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<Control-c>", self.copy_to_clipboard)
        self.tree.bind("<Control-v>", self.paste_from_clipboard)

        # ==================== НИЖНЯЯ ПАНЕЛЬ ====================
        bottom_frame = tk.Frame(root, relief="sunken", bd=1)
        bottom_frame.pack(fill="x", padx=10, pady=8)

        tk.Button(bottom_frame, text="💾 Сохранить все изменения", width=25, height=2,
                 bg="#2E7D32", fg="white", font=("Arial", 10, "bold"),
                 command=lambda: self.save_data(show_msg=True)).pack(side="left", padx=10)

        tk.Button(bottom_frame, text="🚪 Выход", width=15, height=2, bg="#f44336", fg="white",
                 command=self.on_close).pack(side="right", padx=10)

        self.refresh_tree()

    # ==================== ОКНО ДОБАВЛЕНИЯ НОВОГО МАТЕРИАЛА ====================
    def open_add_material_window(self):
        win = tk.Toplevel(self.root)
        win.title("Добавление нового материала")
        win.geometry("700x600")
        win.resizable(True, True)

        fields = [
            ("Производитель", "manufacturer"),
            ("Вид материала", "material_type"),
            ("Паспорт №", "passport_num"),
            ("Дата производства", "production_date"),
            ("Срок хранения", "shelf_life"),
            ("Сертификат №", "cert_num"),
            ("Дата выдачи сертификата", "cert_issue_date"),
            ("Дата окончания сертификата", "cert_exp_date"),
            ("Протокол №", "lab_protocol_num"),
            ("Дата протокола", "lab_protocol_date"),
            ("Акт отбора №", "sample_act_num"),
            ("Дата акта", "sample_act_date"),
        ]

        entries = {}
        for i, (label_text, key) in enumerate(fields):
            tk.Label(win, text=label_text + ":", font=("Arial", 10)).grid(row=i, column=0, sticky="e", padx=10, pady=6)
            entry = tk.Entry(win, width=60, font=("Arial", 10))
            entry.grid(row=i, column=1, padx=10, pady=6)
            entries[key] = entry

        def save_material():
            new_item = {"id": self.next_id}
            self.next_id += 1

            for key, entry in entries.items():
                new_item[key] = entry.get().strip()

            self.data.append(new_item)
            self.refresh_tree()
            self.save_data()
            win.destroy()
            messagebox.showinfo("Успешно", "Новый материал добавлен!")

        tk.Button(win, text="✅ Сохранить материал", width=20, height=2, bg="#4CAF50", fg="white",
                 command=save_material).pack(pady=20)

        tk.Button(win, text="Добавить дополнительное поле", bg="#FF9800", fg="white",
                 command=self.add_new_column_from_window).pack(pady=5)

    def add_new_column_from_window(self):
        name = simpledialog.askstring("Новое поле", 
            "Введите название нового поля\n(например: Партия, Количество, Примечание, Поставщик)")
        if not name:
            return
        key = name.strip().lower().replace(" ", "_").replace(":", "_")
        
        if key in self.tree["columns"]:
            messagebox.showwarning("Ошибка", "Такое поле уже существует!")
            return

        for item in self.data:
            item[key] = ""

        self.refresh_columns()
        self.refresh_tree()
        messagebox.showinfo("Добавлено", f"Поле «{name}» добавлено во все записи.")

    # ==================== Остальные методы (сокращённо) ====================
    def get_data_file(self):
        # (тот же код, что был раньше)
        config_file = "config.json"
        if os.path.exists(config_file):
            try:
                with open(config_file, "r", encoding="utf-8") as f:
                    config = json.load(f)
                    if config.get("data_file") and os.path.exists(config["data_file"]):
                        return config["data_file"]
            except:
                pass

        file_path = filedialog.askopenfilename(title="Выберите materials.json", 
                                              filetypes=[("JSON файл", "materials.json")])
        if not file_path:
            self.root.destroy()
            exit()

        with open(config_file, "w", encoding="utf-8") as f:
            json.dump({"data_file": file_path}, f, ensure_ascii=False, indent=2)
        return file_path

    # ... (refresh_columns, load_data, save_data, refresh_tree, filter_data и т.д. — оставляем как в предыдущей версии)

    def refresh_columns(self):
        if not self.data:
            cols = ["manufacturer", "material_type", "cert_exp_date"]
        else:
            cols = [k for k in self.data[0].keys() if k != "id"]

        self.tree["columns"] = cols
        russian_names = {
            "manufacturer": "Производитель", "material_type": "Вид материала",
            "passport_num": "Паспорт №", "production_date": "Дата производства",
            "shelf_life": "Срок хранения", "cert_num": "Сертификат №",
            "cert_issue_date": "Дата выдачи", "cert_exp_date": "Дата окончания сертификата",
            "lab_protocol_num": "Протокол №", "lab_protocol_date": "Дата протокола",
            "sample_act_num": "Акт отбора №", "sample_act_date": "Дата акта"
        }
        for col in cols:
            name = russian_names.get(col, col.replace("_", " ").title())
            self.tree.heading(col, text=name)
            self.tree.column(col, width=160, anchor="center")

    # Добавь остальные методы из предыдущей версии (on_double_click, copy, paste, save_data и т.д.)

    def on_close(self):
        self.save_data()
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = MaterialApp(root)
    root.mainloop()