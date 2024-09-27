import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import json
from tkinter import filedialog

import openpyxl


class Database:
    def __init__(self, db_name='notary_office.db'):
        self.db_name = db_name
        self.setup_database()

    def connect_db(self):
        conn = sqlite3.connect(self.db_name)
        conn.execute("PRAGMA foreign_keys = ON")  # Включение поддержки внешних ключей
        return conn

    def setup_database(self):
        conn = self.connect_db()
        c = conn.cursor()

        c.execute('''
            CREATE TABLE IF NOT EXISTS Клиенты (
                Код_клиента INTEGER PRIMARY KEY,
                Название TEXT,
                Вид_деятельности TEXT,
                Адрес TEXT,
                Телефон TEXT
            )
        ''')
        c.execute('''
            CREATE TABLE IF NOT EXISTS Услуги (
                Код_услуги INTEGER PRIMARY KEY,
                Название TEXT,
                Описание TEXT
            )
        ''')
        c.execute('''
            CREATE TABLE IF NOT EXISTS Сделки (
                Код_сделки INTEGER PRIMARY KEY,
                Код_клиента INTEGER,
                Код_услуги INTEGER,
                Сумма REAL,
                Комиссионные REAL,
                Описание TEXT,
                FOREIGN KEY (Код_клиента) REFERENCES Клиенты (Код_клиента) ON DELETE CASCADE,
                FOREIGN KEY (Код_услуги) REFERENCES Услуги (Код_услуги) ON DELETE CASCADE
            )
        ''')

    def export_to_json(self, table_name):
        conn = self.connect_db()
        c = conn.cursor()
        c.execute(f"SELECT * FROM {table_name}")
        rows = c.fetchall()
        columns = [desc[0] for desc in c.description]
        data = [dict(zip(columns, row)) for row in rows]

        with open(f"{table_name}.json", "w", encoding="utf-8") as file:
            json.dump(data, file, ensure_ascii=False, indent=4)

        # messagebox.showinfo("Экспорт данных", f"Данные успешно экспортированы в {table_name}.json")

        return f"{table_name}.json"

    def import_from_json(self, table_name):

        # Окно выбора файла для импорта

        file_path = filedialog.askopenfilename(
            title="Выберите файл для импорта",
            filetypes=[("JSON файлы", "*.json")]
        )
        if not file_path:
            messagebox.showwarning("Ошибка", "Файл не выбран.")
            return

        try:
            with open(file_path, "r", encoding="utf-8") as file:
                data = json.load(file)
        except FileNotFoundError:
            raise FileNotFoundError(f"{table_name}.json not found.")
        except json.JSONDecodeError:
            raise ValueError(f"Error decoding JSON from {table_name}.json")

        conn = self.connect_db()
        c = conn.cursor()
        for row in data:
            columns = ', '.join(row.keys())
            placeholders = ', '.join('?' * len(row))
            values = tuple(row.values())
            c.execute(f"INSERT OR REPLACE INTO {table_name} ({columns}) VALUES ({placeholders})", values)
        conn.commit()
        conn.close()

    def export_to_excel(self, table_name):
        # Создание новой рабочей книги и рабочего листа
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = table_name

        conn = self.connect_db()
        c = conn.cursor()
        c.execute(f"SELECT * FROM {table_name}")
        rows = c.fetchall()
        columns = [desc[0] for desc in c.description]

        # Запись заголовков столбцов
        sheet.append(columns)

        # Запись данных строк
        for row in rows:
            sheet.append(row)

        file_path = f"{table_name}.xlsx"
        workbook.save(file_path)

        return file_path


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Система управления нотариальной конторой")
        self.db = Database()
        self.create_widgets()

    def create_widgets(self):
        # Создание менюбара
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Файл", menu=file_menu)
        file_menu.add_command(label="Экспорт Клиенты", command=lambda: self.export_data('Клиенты'))
        file_menu.add_command(label="Экспорт Услуги", command=lambda: self.export_data('Услуги'))
        file_menu.add_command(label="Экспорт Сделки", command=lambda: self.export_data('Сделки'))
        file_menu.add_command(label="Импорт Клиенты", command=lambda: self.import_data('Клиенты'))
        file_menu.add_command(label="Импорт Услуги", command=lambda: self.import_data('Услуги'))
        file_menu.add_command(label="Импорт Сделки", command=lambda: self.import_data('Сделки'))
        file_menu.add_command(label="Экспорт Клиенты в Excel",
                              command=lambda: self.export_data_excel('Клиенты', excel=True))
        file_menu.add_command(label="Экспорт Услуги в Excel",
                              command=lambda: self.export_data_excel('Услуги', excel=True))
        file_menu.add_command(label="Экспорт Сделки в Excel",
                              command=lambda: self.export_data_excel('Сделки', excel=True))

        frame = ttk.Frame(self.root)
        frame.grid(row=0, column=0, padx=10, pady=10, sticky='nsew')

        tab_control = ttk.Notebook(frame)
        tab_clients = ttk.Frame(tab_control)
        tab_services = ttk.Frame(tab_control)
        tab_transactions = ttk.Frame(tab_control)

        tab_control.add(tab_clients, text='Клиенты')
        tab_control.add(tab_services, text='Услуги')
        tab_control.add(tab_transactions, text='Сделки')
        tab_control.grid(row=0, column=0, padx=5, pady=5, sticky='nsew')

        self.create_client_tab(tab_clients)
        self.create_service_tab(tab_services)
        self.create_transaction_tab(tab_transactions)

    def export_data_excel(self, table_name, excel=False):
        try:
            if excel:
                file_path = self.db.export_to_excel(table_name)
                messagebox.showinfo("Экспорт данных", f"Данные успешно экспортированы в {file_path}.")
            else:
                file_path = self.db.export_to_json(table_name)
                messagebox.showinfo("Экспорт данных", f"Данные успешно экспортированы в {file_path}.")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def export_data(self, table_name):
        try:
            file_path = self.db.export_to_json(table_name)
            messagebox.showinfo("Экспорт данных", f"Данные успешно экспортированы в {file_path}.")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def import_data(self, table_name):
        try:
            self.db.import_from_json(table_name)
            messagebox.showinfo("Импорт данных", "Данные успешно импортированы.")
            if table_name == "Клиенты":
                self.update_treeview(self.client_tree, "SELECT * FROM Клиенты",
                                     ("ID", "Название", "Вид деятельности", "Адрес", "Телефон"))
            elif table_name == "Услуги":
                self.update_treeview(self.service_tree, "SELECT * FROM Услуги", ("ID", "Название", "Описание"))
            elif table_name == "Сделки":
                self.update_treeview(self.transaction_tree, "SELECT * FROM Сделки",
                                     ("ID", "Код клиента", "Код услуги", "Сумма", "Комиссионные", "Описание"))
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def update_treeview(self, treeview, query, columns):
        conn = self.db.connect_db()
        c = conn.cursor()
        c.execute(query)
        rows = c.fetchall()
        treeview.delete(*treeview.get_children())
        for row in rows:
            treeview.insert('', 'end', values=row)

    def create_client_tab(self, parent):
        client_frame = ttk.Frame(parent)
        client_frame.pack(padx=10, pady=10, fill='both', expand=True)

        ttk.Label(client_frame, text="ID клиента").grid(row=0, column=0, padx=5, pady=5)
        self.client_id = tk.StringVar()
        ttk.Entry(client_frame, textvariable=self.client_id).grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(client_frame, text="Название").grid(row=1, column=0, padx=5, pady=5)
        self.client_name = tk.StringVar()
        ttk.Entry(client_frame, textvariable=self.client_name).grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(client_frame, text="Вид деятельности").grid(row=2, column=0, padx=5, pady=5)
        self.client_activity = tk.StringVar()
        ttk.Entry(client_frame, textvariable=self.client_activity).grid(row=2, column=1, padx=5, pady=5)

        ttk.Label(client_frame, text="Адрес").grid(row=3, column=0, padx=5, pady=5)
        self.client_address = tk.StringVar()
        ttk.Entry(client_frame, textvariable=self.client_address).grid(row=3, column=1, padx=5, pady=5)

        ttk.Label(client_frame, text="Телефон").grid(row=4, column=0, padx=5, pady=5)
        self.client_phone = tk.StringVar()
        ttk.Entry(client_frame, textvariable=self.client_phone).grid(row=4, column=1, padx=5, pady=5)

        ttk.Button(client_frame, text="Добавить", command=self.add_client).grid(row=5, column=0, padx=5, pady=5)
        ttk.Button(client_frame, text="Обновить", command=self.update_client).grid(row=5, column=1, padx=5, pady=5)
        ttk.Button(client_frame, text="Удалить", command=self.delete_client).grid(row=5, column=2, padx=5, pady=5)

        self.client_tree = ttk.Treeview(parent, columns=("ID", "Название", "Вид деятельности", "Адрес", "Телефон"),
                                        show='headings')
        self.client_tree.heading("ID", text="ID")
        self.client_tree.heading("Название", text="Название")
        self.client_tree.heading("Вид деятельности", text="Вид деятельности")
        self.client_tree.heading("Адрес", text="Адрес")
        self.client_tree.heading("Телефон", text="Телефон")
        self.client_tree.pack(fill='both', expand=True)

        ttk.Label(client_frame, text="Поиск по ID клиента").grid(row=6, column=0, padx=5, pady=5)
        search_entry = tk.StringVar()
        self.client_entry = ttk.Entry(client_frame, textvariable=search_entry)
        self.client_entry.grid(row=6, column=1, padx=5, pady=5)

        ttk.Button(client_frame, text="Искать",
                   command=lambda: self.search_data(self.client_entry, self.client_tree,
                                                    ("Клиенты", "Код_клиента"))).grid(row=6,
                                                                                      column=2,
                                                                                      padx=5,
                                                                                      pady=5)

        self.update_treeview(self.client_tree, "SELECT * FROM Клиенты",
                             ("ID", "Название", "Вид деятельности", "Адрес", "Телефон"))

    def create_service_tab(self, parent):
        service_frame = ttk.Frame(parent)
        service_frame.pack(padx=10, pady=10, fill='both', expand=True)

        ttk.Label(service_frame, text="ID услуги").grid(row=0, column=0, padx=5, pady=5)
        self.service_id = tk.StringVar()
        ttk.Entry(service_frame, textvariable=self.service_id).grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(service_frame, text="Название").grid(row=1, column=0, padx=5, pady=5)
        self.service_name = tk.StringVar()
        ttk.Entry(service_frame, textvariable=self.service_name).grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(service_frame, text="Описание").grid(row=2, column=0, padx=5, pady=5)
        self.service_description = tk.StringVar()
        ttk.Entry(service_frame, textvariable=self.service_description).grid(row=2, column=1, padx=5, pady=5)

        ttk.Button(service_frame, text="Добавить", command=self.add_service).grid(row=3, column=0, padx=5, pady=5)
        ttk.Button(service_frame, text="Обновить", command=self.update_service).grid(row=3, column=1, padx=5, pady=5)
        ttk.Button(service_frame, text="Удалить", command=self.delete_service).grid(row=3, column=2, padx=5, pady=5)

        self.service_tree = ttk.Treeview(parent, columns=("ID", "Название", "Описание"), show='headings')
        self.service_tree.heading("ID", text="ID")
        self.service_tree.heading("Название", text="Название")
        self.service_tree.heading("Описание", text="Описание")
        self.service_tree.pack(fill='both', expand=True)

        ttk.Label(service_frame, text="Поиск по ID услуги").grid(row=4, column=0, padx=5, pady=5)
        search_entry = tk.StringVar()
        self.service_entry = ttk.Entry(service_frame, textvariable=search_entry)
        self.service_entry.grid(row=4, column=1, padx=5, pady=5)

        ttk.Button(service_frame, text="Искать",
                   command=lambda: self.search_data(self.service_entry, self.service_tree,
                                                    ("Услуги", "Код_услуги"))).grid(row=4,
                                                                                    column=2,
                                                                                    padx=5,
                                                                                    pady=5)

        self.update_treeview(self.service_tree, "SELECT * FROM Услуги", ("ID", "Название", "Описание"))

    def create_transaction_tab(self, parent):
        transaction_frame = ttk.Frame(parent)
        transaction_frame.pack(padx=10, pady=10, fill='both', expand=True)

        ttk.Label(transaction_frame, text="ID сделки").grid(row=0, column=0, padx=5, pady=5)
        self.transaction_id = tk.StringVar()
        ttk.Entry(transaction_frame, textvariable=self.transaction_id).grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(transaction_frame, text="ID клиента").grid(row=1, column=0, padx=5, pady=5)
        self.transaction_client_id = tk.StringVar()
        ttk.Entry(transaction_frame, textvariable=self.transaction_client_id).grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(transaction_frame, text="ID услуги").grid(row=2, column=0, padx=5, pady=5)
        self.transaction_service_id = tk.StringVar()
        ttk.Entry(transaction_frame, textvariable=self.transaction_service_id).grid(row=2, column=1, padx=5, pady=5)

        ttk.Label(transaction_frame, text="Сумма").grid(row=3, column=0, padx=5, pady=5)
        self.transaction_amount = tk.StringVar()
        ttk.Entry(transaction_frame, textvariable=self.transaction_amount).grid(row=3, column=1, padx=5, pady=5)

        ttk.Label(transaction_frame, text="Комиссионные").grid(row=4, column=0, padx=5, pady=5)
        self.transaction_commission = tk.StringVar()
        ttk.Entry(transaction_frame, textvariable=self.transaction_commission).grid(row=4, column=1, padx=5, pady=5)

        ttk.Label(transaction_frame, text="Описание").grid(row=5, column=0, padx=5, pady=5)
        self.transaction_description = tk.StringVar()
        ttk.Entry(transaction_frame, textvariable=self.transaction_description).grid(row=5, column=1, padx=5, pady=5)

        ttk.Button(transaction_frame, text="Добавить", command=self.add_transaction).grid(row=6, column=0, padx=5,
                                                                                          pady=5)
        ttk.Button(transaction_frame, text="Обновить", command=self.update_transaction).grid(row=6, column=1, padx=5,
                                                                                             pady=5)
        ttk.Button(transaction_frame, text="Удалить", command=self.delete_transaction).grid(row=6, column=2, padx=5,
                                                                                            pady=5)

        self.transaction_tree = ttk.Treeview(parent, columns=(
            "ID", "Код клиента", "Код услуги", "Сумма", "Комиссионные", "Описание"), show='headings')
        self.transaction_tree.heading("ID", text="ID")
        self.transaction_tree.heading("Код клиента", text="Код клиента")
        self.transaction_tree.heading("Код услуги", text="Код услуги")
        self.transaction_tree.heading("Сумма", text="Сумма")
        self.transaction_tree.heading("Комиссионные", text="Комиссионные")
        self.transaction_tree.heading("Описание", text="Описание")
        self.transaction_tree.pack(fill='both', expand=True)

        # Поиск

        ttk.Label(transaction_frame, text="Поиск по ID сделки").grid(row=7, column=0, padx=5, pady=5)
        search_entry = tk.StringVar()
        self.transaction_entry = ttk.Entry(transaction_frame, textvariable=search_entry)
        self.transaction_entry.grid(row=7, column=1, padx=5, pady=5)

        ttk.Button(transaction_frame, text="Искать",
                   command=lambda: self.search_data(self.transaction_entry, self.transaction_tree,
                                                    ("Сделки", "Код_сделки"))).grid(
            row=7,
            column=2,
            padx=5,
            pady=5)

        self.update_treeview(self.transaction_tree, "SELECT * FROM Сделки",
                             ("ID", "Код клиента", "Код услуги", "Сумма", "Комиссионные", "Описание"))

    def search_data(self, entry, client_tree, where):
        search_id = entry.get()
        if not search_id:
            messagebox.showerror("Error", "Введите ID для поиска.")
            return

        try:
            search_id = int(search_id)
        except ValueError:
            messagebox.showerror("Error", "Некорректный ID.")
            return
        conn = self.db.connect_db()
        c = conn.cursor()

        c.execute(f'SELECT * FROM {where[0]} WHERE {where[1]} = ?', (search_id,))

        result = c.fetchone()
        conn.close()

        if result:
            # Очистка предыдущих выделений
            for item in client_tree.selection():
                client_tree.selection_remove(item)

            # Поиск строки и выделение её
            for child in client_tree.get_children():
                if client_tree.item(child)['values'][0] == result[0]:  # Сравнение по Код_клиента
                    client_tree.selection_set(child)
                    client_tree.see(child)  # Прокрутка к выделенной строке
                    break
        else:
            messagebox.showinfo("Info", "ID не найден.")

    def add_client(self):
        conn = self.db.connect_db()
        c = conn.cursor()
        c.execute('''
            INSERT OR REPLACE INTO Клиенты (Код_клиента, Название, Вид_деятельности, Адрес, Телефон)
            VALUES (?, ?, ?, ?, ?)
        ''', (self.client_id.get(),
              self.client_name.get(), self.client_activity.get(), self.client_address.get(),
              self.client_phone.get()))
        conn.commit()
        conn.close()
        messagebox.showinfo("Добавление клиента", "Клиент успешно добавлен!")
        self.update_treeview(self.client_tree, "SELECT * FROM Клиенты",
                             ("ID", "Название", "Вид деятельности", "Адрес", "Телефон"))

    def update_client(self):
        conn = self.db.connect_db()
        c = conn.cursor()
        c.execute('''
            UPDATE Клиенты
            SET Название = ?, Вид_деятельности = ?, Адрес = ?, Телефон = ?
            WHERE Код_клиента = ?
        ''', (
            self.client_name.get(), self.client_activity.get(), self.client_address.get(), self.client_phone.get(),
            self.client_id.get()))
        conn.commit()
        conn.close()
        messagebox.showinfo("Обновление клиента", "Клиент успешно обновлён!")
        self.update_treeview(self.client_tree, "SELECT * FROM Клиенты",
                             ("ID", "Название", "Вид деятельности", "Адрес", "Телефон"))

    def delete_client(self):
        conn = self.db.connect_db()
        c = conn.cursor()
        c.execute('DELETE FROM Клиенты WHERE Код_клиента = ?', (int(self.client_id.get()),))
        conn.commit()
        conn.close()
        messagebox.showinfo("Удаление клиента", "Клиент успешно удалён!")
        self.update_treeview(self.client_tree, "SELECT * FROM Клиенты",
                             ("ID", "Название", "Вид деятельности", "Адрес", "Телефон"))

        self.update_treeview(self.transaction_tree, "SELECT * FROM Сделки",
                             ("ID", "Код клиента", "Код услуги", "Сумма", "Комиссионные", "Описание"))

    def add_service(self):
        conn = self.db.connect_db()
        conn.execute('PRAGMA foreign_keys = ON;')
        c = conn.cursor()
        c.execute('''
            INSERT OR REPLACE INTO Услуги (Код_услуги,Название, Описание)
            VALUES (?, ?, ?)
        ''', (self.service_id.get(), self.service_name.get(), self.service_description.get()))
        conn.commit()
        conn.close()
        messagebox.showinfo("Добавление услуги", "Услуга успешно добавлена!")
        self.update_treeview(self.service_tree, "SELECT * FROM Услуги", ("ID", "Название", "Описание"))

    def update_service(self):
        conn = self.db.connect_db()
        c = conn.cursor()
        c.execute('''
            UPDATE Услуги
            SET Название = ?, Описание = ?
            WHERE Код_услуги = ?
        ''', (self.service_name.get(), self.service_description.get(), self.service_id.get()))
        conn.commit()
        conn.close()
        messagebox.showinfo("Обновление услуги", "Услуга успешно обновлена!")
        self.update_treeview(self.service_tree, "SELECT * FROM Услуги", ("ID", "Название", "Описание"))

    def delete_service(self):
        conn = self.db.connect_db()
        conn.execute('PRAGMA foreign_keys = ON;')
        c = conn.cursor()
        c.execute('DELETE FROM Услуги WHERE Код_услуги = ?', (int(self.service_id.get()),))
        conn.commit()
        conn.close()
        messagebox.showinfo("Удаление услуги", "Услуга успешно удалена!")
        self.update_treeview(self.service_tree, "SELECT * FROM Услуги", ("ID", "Название", "Описание"))
        self.update_treeview(self.transaction_tree, "SELECT * FROM Сделки",
                             ("ID", "Код клиента", "Код услуги", "Сумма", "Комиссионные", "Описание"))

    def add_transaction(self):
        conn = self.db.connect_db()
        c = conn.cursor()
        c.execute('''
            INSERT OR REPLACE INTO Сделки (Код_сделки, Код_клиента, Код_услуги, Сумма, Комиссионные, Описание)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (self.transaction_id.get(), self.transaction_client_id.get(), self.transaction_service_id.get(),
              self.transaction_amount.get(),
              self.transaction_commission.get(), self.transaction_description.get()))
        conn.commit()
        conn.close()
        messagebox.showinfo("Добавление сделки", "Сделка успешно добавлена!")
        self.update_treeview(self.transaction_tree, "SELECT * FROM Сделки",
                             ("ID", "Код клиента", "Код услуги", "Сумма", "Комиссионные", "Описание"))

    def update_transaction(self):
        conn = self.db.connect_db()
        c = conn.cursor()
        c.execute('''
            UPDATE Сделки
            SET Код_клиента = ?, Код_услуги = ?, Сумма = ?, Комиссионные = ?, Описание = ?
            WHERE Код_сделки = ?
        ''', (self.transaction_client_id.get(), self.transaction_service_id.get(), self.transaction_amount.get(),
              self.transaction_commission.get(), self.transaction_description.get(), self.transaction_id.get()))
        conn.commit()
        conn.close()
        messagebox.showinfo("Обновление сделки", "Сделка успешно обновлена!")
        self.update_treeview(self.transaction_tree, "SELECT * FROM Сделки",
                             ("ID", "Код клиента", "Код услуги", "Сумма", "Комиссионные", "Описание"))

    def delete_transaction(self):
        conn = self.db.connect_db()
        conn.execute('PRAGMA foreign_keys = ON;')
        c = conn.cursor()
        c.execute('DELETE FROM Сделки WHERE Код_сделки = ?', (int(self.transaction_id.get()),))
        conn.commit()
        conn.close()
        messagebox.showinfo("Удаление сделки", "Сделка успешно удалена!")
        self.update_treeview(self.transaction_tree, "SELECT * FROM Сделки",
                             ("ID", "Код клиента", "Код услуги", "Сумма", "Комиссионные", "Описание"))


if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = App(root)
        root.mainloop()
    except KeyboardInterrupt:
        print("Program interrupted by user")
        quit()
