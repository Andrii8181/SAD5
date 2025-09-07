import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
from stat_analysis import StatAnalysis
from report_generator import ReportGenerator

class SADApp(tk.Tk):
    """
    Головний клас програми SAD (Статистичний аналіз даних).
    Відповідає за створення та управління інтерфейсом користувача.
    """
    def __init__(self):
        super().__init__()
        self.title("SAD - Статистичний аналіз даних")
        self.geometry("1000x700")

        self.df = pd.DataFrame(np.zeros((10, 10)), dtype=object) # Змінено dtype на object для зберігання тексту
        self.create_widgets()
        self.selected_item = None
        self.selected_col_index = -1

    def create_widgets(self):
        """Створює всі елементи інтерфейсу: кнопки та таблицю."""
        control_frame = ttk.Frame(self)
        control_frame.pack(fill='x', padx=10, pady=5)

        ttk.Button(control_frame, text="Додати рядок", command=self.add_row).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Видалити рядок", command=self.remove_row).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Додати стовпець", command=self.add_column).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Видалити стовпець", command=self.remove_column).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Аналіз даних", command=self.analyze_data).pack(side='right', padx=5)

        self.tree = ttk.Treeview(self, show="headings")
        self.tree.pack(expand=True, fill='both', padx=10, pady=5)
        self.update_table()
        self.tree.bind('<Button-3>', self.show_context_menu)
        self.tree.bind('<Double-1>', self.edit_cell)
        self.tree.bind('<Button-1>', self.select_cell)

    def update_table(self):
        """Оновлює дані в таблиці Treeview з DataFrame."""
        self.tree.delete(*self.tree.get_children())
        cols = list(self.df.columns)
        self.tree["columns"] = cols
        for i, col in enumerate(cols):
            self.tree.heading(col, text=f"Col {i+1}")
            self.tree.column(col, width=100)

        self.tree.heading("#0", text="Рядок")
        self.tree.column("#0", width=80)

        for i, row in self.df.iterrows():
            self.tree.insert("", "end", text=f"Row {i+1}", values=list(row))

    def add_row(self):
        """Додає новий рядок до таблиці."""
        new_row = pd.Series([0] * len(self.df.columns), dtype=object)
        self.df = pd.concat([self.df, pd.DataFrame([new_row])], ignore_index=True)
        self.update_table()

    def remove_row(self):
        """Видаляє останній рядок з таблиці."""
        if len(self.df) > 1:
            self.df = self.df.iloc[:-1, :]
            self.update_table()

    def add_column(self):
        """Додає новий стовпець до таблиці."""
        new_col_name = len(self.df.columns)
        self.df[new_col_name] = [0] * len(self.df)
        self.update_table()

    def remove_column(self):
        """Видаляє останній стовпець з таблиці."""
        if len(self.df.columns) > 1:
            self.df = self.df.iloc[:, :-1]
            self.update_table()
    
    def show_context_menu(self, event):
        """Створює контекстне меню для таблиці."""
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="Вставити дані", command=self.paste_data)
        menu.post(event.x_root, event.y_root)

    def paste_data(self):
        """Вставляє дані з буфера обміну (включаючи Excel)."""
        try:
            clipboard_content = self.clipboard_get()
            rows = clipboard_content.strip().split('\n')
            pasted_data = [row.split('\t') for row in rows]
            pasted_df = pd.DataFrame(pasted_data)
            
            self.df.iloc[:pasted_df.shape[0], :pasted_df.shape[1]] = pasted_df.values
            self.update_table()
            messagebox.showinfo("Успіх", "Дані успішно вставлено!")
        except Exception as e:
            messagebox.showerror("Помилка", f"Не вдалося вставити дані: {e}")

    def select_cell(self, event):
        """Отримує індекс вибраної клітинки."""
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if item and col:
            self.selected_item = item
            self.selected_col_index = int(col.replace("#", "")) - 1
            self.tree.selection_remove(self.tree.selection())
            self.tree.selection_add(self.selected_item)

    def edit_cell(self, event):
        """Дозволяє редагувати дані в клітинці."""
        item_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not item_id or not col_id:
            return

        row_index = int(self.tree.index(item_id))
        col_index = int(col_id.replace("#", "")) - 1
        
        cell_value = self.df.iloc[row_index, col_index]
        entry = ttk.Entry(self, width=10)
        entry.insert(0, str(cell_value))
        entry.bind("<Return>", lambda e: self.save_edit(e, entry, row_index, col_index))
        entry.bind("<FocusOut>", lambda e: self.save_edit(e, entry, row_index, col_index, focus_out=True))
        
        # Key bindings for navigation
        entry.bind('<Right>', lambda e: self.move_focus(1, 0, entry))
        entry.bind('<Left>', lambda e: self.move_focus(-1, 0, entry))
        entry.bind('<Up>', lambda e: self.move_focus(0, -1, entry))
        entry.bind('<Down>', lambda e: self.move_focus(0, 1, entry))
        entry.bind('<Tab>', lambda e: self.move_focus(1, 0, entry, tab=True))

        x, y, width, height = self.tree.bbox(item_id, col_id)
        entry.place(x=x, y=y, width=width, height=height)
        entry.focus_set()
        entry.select_range(0, 'end')

    def save_edit(self, event, entry, row_index, col_index, focus_out=False):
        """Зберігає змінене значення клітинки."""
        if not entry.winfo_exists():
            return
        
        new_value = entry.get()
        try:
            self.df.iloc[row_index, col_index] = new_value
            self.update_table()
        except ValueError:
            messagebox.showerror("Помилка", "Некоректне значення.")
        finally:
            entry.destroy()
            if not focus_out:
                self.focus()

    def move_focus(self, col_shift, row_shift, current_entry, tab=False):
        """Переміщує фокус на іншу клітинку."""
        current_entry.destroy()
        
        row_index = int(self.tree.index(self.selected_item))
        col_index = self.selected_col_index

        new_row_index = row_index + row_shift
        new_col_index = col_index + col_shift

        # Обробка переходу на наступний рядок при натисканні Tab або Enter
        if tab or (row_shift == 0 and col_shift == 1):
            if new_col_index >= len(self.df.columns):
                new_col_index = 0
                new_row_index += 1
            if new_col_index < 0:
                new_col_index = len(self.df.columns) - 1
                new_row_index -= 1

        # Перевірка меж таблиці
        if 0 <= new_row_index < len(self.df) and 0 <= new_col_index < len(self.df.columns):
            new_item_id = self.tree.get_children()[new_row_index]
            self.tree.selection_remove(self.tree.selection())
            self.tree.selection_add(new_item_id)
            
            self.selected_item = new_item_id
            self.selected_col_index = new_col_index

            # Запускаємо редагування нової клітинки
            self.edit_cell_by_indices(new_row_index, new_col_index)
        else:
            self.focus()

    def edit_cell_by_indices(self, row_index, col_index):
        """Запускає редагування клітинки за індексами."""
        item_id = self.tree.get_children()[row_index]
        col_id = f"#{col_index + 1}"
        
        x, y, width, height = self.tree.bbox(item_id, col_id)
        
        cell_value = self.df.iloc[row_index, col_index]
        entry = ttk.Entry(self, width=10)
        entry.insert(0, str(cell_value))
        entry.bind("<Return>", lambda e: self.save_edit(e, entry, row_index, col_index))
        entry.bind("<FocusOut>", lambda e: self.save_edit(e, entry, row_index, col_index, focus_out=True))
        
        entry.bind('<Right>', lambda e: self.move_focus(1, 0, entry))
        entry.bind('<Left>', lambda e: self.move_focus(-1, 0, entry))
        entry.bind('<Up>', lambda e: self.move_focus(0, -1, entry))
        entry.bind('<Down>', lambda e: self.move_focus(0, 1, entry))
        entry.bind('<Tab>', lambda e: self.move_focus(1, 0, entry, tab=True))

        entry.place(x=x, y=y, width=width, height=height)
        entry.focus_set()
        entry.select_range(0, 'end')


    def analyze_data(self):
        """Запускає процес аналізу даних."""
        StatAnalysis(self.df).run_analysis(self)

if __name__ == "__main__":
    app = SADApp()
    app.mainloop()
