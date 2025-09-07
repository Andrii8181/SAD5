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

        self.df = pd.DataFrame(np.zeros((10, 10)), dtype=object)
        self.create_widgets()
        self.selected_item = None
        self.selected_col_index = -1
        self.active_entry = None

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
        self.tree.bind('<Button-1>', self.select_cell)
        
        # New bindings for direct editing and navigation
        self.tree.bind('<Return>', lambda e: self.edit_cell(e))
        self.tree.bind('<Tab>', lambda e: self.edit_cell(e))
        self.tree.bind('<Left>', lambda e: self.move_focus(-1, 0, e))
        self.tree.bind('<Right>', lambda e: self.move_focus(1, 0, e))
        self.tree.bind('<Up>', lambda e: self.move_focus(0, -1, e))
        self.tree.bind('<Down>', lambda e: self.move_focus(0, 1, e))
        self.tree.bind('<Key>', self.start_edit_on_key)
        
        # Bindings for copy/paste
        self.bind_all("<Control-c>", self.copy_data)
        self.bind_all("<Control-v>", self.paste_data)


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

    def paste_data(self, event=None):
        """Вставляє дані з буфера обміну (включаючи Excel)."""
        try:
            clipboard_content = self.clipboard_get()
            rows = clipboard_content.strip().split('\n')
            pasted_data = [row.split('\t') for row in rows]
            pasted_df = pd.DataFrame(pasted_data)
            
            # Вставка в поточну виділену клітинку
            if self.selected_item:
                row_index = int(self.tree.index(self.selected_item))
                col_index = self.selected_col_index
                
                self.df.iloc[row_index : row_index + pasted_df.shape[0], 
                             col_index : col_index + pasted_df.shape[1]] = pasted_df.values
            else:
                self.df.iloc[:pasted_df.shape[0], :pasted_df.shape[1]] = pasted_df.values

            self.update_table()
            messagebox.showinfo("Успіх", "Дані успішно вставлено!")
        except Exception as e:
            messagebox.showerror("Помилка", f"Не вдалося вставити дані: {e}")

    def copy_data(self, event=None):
        """Копіює дані з виділеної клітинки в буфер обміну."""
        if self.selected_item:
            try:
                row_index = int(self.tree.index(self.selected_item))
                col_index = self.selected_col_index
                value = self.df.iloc[row_index, col_index]
                self.clipboard_clear()
                self.clipboard_append(str(value))
            except Exception as e:
                pass

    def select_cell(self, event):
        """Отримує індекс вибраної клітинки."""
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if item and col:
            self.selected_item = item
            self.selected_col_index = int(col.replace("#", "")) - 1
            self.tree.selection_remove(self.tree.selection())
            self.tree.selection_add(self.selected_item)

    def start_edit_on_key(self, event):
        """Починає редагування при натисканні будь-якої клавіші."""
        if self.selected_item and event.char.isalnum():
            self.edit_cell(event, start_empty=True)

    def edit_cell(self, event=None, start_empty=False):
        """Запускає редагування поточної вибраної клітинки."""
        if self.active_entry:
            return

        item_id = self.tree.identify_row(event.y) if event else self.selected_item
        col_id = self.tree.identify_column(event.x) if event else f"#{self.selected_col_index + 1}"
        
        if not item_id or not col_id:
            return

        row_index = int(self.tree.index(item_id))
        col_index = int(col_id.replace("#", "")) - 1
        
        cell_value = "" if start_empty else self.df.iloc[row_index, col_index]
        
        entry = ttk.Entry(self, width=10)
        entry.insert(0, str(cell_value))

        def save_edit(event):
            try:
                new_value = entry.get()
                self.df.iloc[row_index, col_index] = new_value
                self.update_table()
                entry.destroy()
                self.active_entry = None
            except Exception as e:
                messagebox.showerror("Помилка", f"Помилка введення: {e}")
                entry.destroy()
                self.active_entry = None
                self.tree.focus_set()

        def on_return(event):
            save_edit(event)
            self.move_focus(1, 0, event)

        entry.bind("<Return>", on_return)
        entry.bind("<Tab>", on_return)
        entry.bind("<FocusOut>", save_edit)

        x, y, width, height = self.tree.bbox(item_id, self.tree["columns"][col_index])
        entry.place(x=x, y=y, width=width, height=height)
        entry.focus_set()
        entry.select_range(0, 'end')
        self.active_entry = entry

    def move_focus(self, col_shift, row_shift, event=None):
        """Переміщує фокус на іншу клітинку."""
        if self.active_entry:
            self.active_entry.destroy()
            self.active_entry = None

        if not self.selected_item:
            return
        
        row_index = int(self.tree.index(self.selected_item))
        col_index = self.selected_col_index

        new_row_index = row_index + row_shift
        new_col_index = col_index + col_shift

        if event and event.keysym in ['Return', 'Tab']:
            new_col_index = col_index + 1
            if new_col_index >= len(self.df.columns):
                new_col_index = 0
                new_row_index += 1

        if 0 <= new_row_index < len(self.df) and 0 <= new_col_index < len(self.df.columns):
            new_item_id = self.tree.get_children()[new_row_index]
            self.tree.selection_remove(self.tree.selection())
            self.tree.selection_add(new_item_id)
            self.selected_item = new_item_id
            self.selected_col_index = new_col_index
            self.tree.focus(new_item_id)
            self.edit_cell_by_indices(new_row_index, new_col_index)
        else:
            self.tree.focus_set()

    def edit_cell_by_indices(self, row_index, col_index):
        """Запускає редагування клітинки за індексами."""
        self.tree.selection_remove(self.tree.selection())
        new_item_id = self.tree.get_children()[row_index]
        self.tree.selection_set(new_item_id)
        self.selected_item = new_item_id
        self.selected_col_index = col_index
        self.edit_cell()

    def analyze_data(self):
        """Запускає процес аналізу даних."""
        StatAnalysis(self.df).run_analysis(self)

if __name__ == "__main__":
    app = SADApp()
    app.mainloop()
