import tkinter as tk
from tkinter import messagebox, simpledialog
from scipy import stats
import statsmodels.api as sm
from statsmodels.formula.api import ols
import numpy as np
import pandas as pd
from report_generator import ReportGenerator

class StatAnalysis:
    """
    Клас для виконання статистичних аналізів.
    """
    def __init__(self, data_frame):
        self.df = data_frame

    def run_analysis(self, parent_app):
        """Виконує тест Шапіро-Вілка та пропонує подальші аналізи."""
        try:
            # Спроба отримати числовий стовпець для тесту
            numeric_cols = self.df.apply(pd.to_numeric, errors='coerce').select_dtypes(include=np.number).columns
            if numeric_cols.empty:
                messagebox.showerror("Помилка", "В таблиці немає числових даних для аналізу.")
                return

            first_numeric_col_name = numeric_cols[0]
            data_to_analyze = self.df[first_numeric_col_name].dropna()
            
            if len(data_to_analyze) < 3:
                messagebox.showerror("Помилка", "Недостатньо даних для аналізу. Потрібно щонайменше 3 значення.")
                return

            shapiro_test = stats.shapiro(data_to_analyze)
            p_value = shapiro_test.pvalue

            if p_value > 0.05:
                messagebox.showinfo("Результат", f"Дані є нормально розподіленими (p={p_value:.3f} > 0.05).")
                self.show_parametric_options(parent_app)
            else:
                messagebox.showinfo("Результат", f"Дані не є нормально розподіленими (p={p_value:.3f} <= 0.05).")
                self.show_non_parametric_options(parent_app)
        except Exception as e:
            messagebox.showerror("Помилка", f"Помилка при аналізі даних: {e}")

    def show_parametric_options(self, parent_app):
        """Відкриває вікно вибору параметричних аналізів."""
        options = ["Дисперсійний аналіз", "Регресія"]
        self.create_analysis_window(parent_app, options, is_parametric=True)

    def show_non_parametric_options(self, parent_app):
        """Відкриває вікно вибору непараметричних аналізів."""
        options = ["Медіанний тест", "Критерій Уїлкоксона"]
        self.create_analysis_window(parent_app, options, is_parametric=False)

    def create_analysis_window(self, parent_app, options, is_parametric):
        """Створює вікно для вибору типу аналізу."""
        analysis_window = tk.Toplevel(parent_app)
        analysis_window.title("Вибір аналізу")
        
        selection = tk.StringVar()
        selection.set(options[0])

        tk.Label(analysis_window, text="Оберіть тип аналізу:").pack(padx=10, pady=5)
        
        for option in options:
            ttk.Radiobutton(analysis_window, text=option, variable=selection, value=option).pack(anchor='w', padx=20)
        
        def run_selected_analysis():
            selected_analysis = selection.get()
            self.perform_and_report_analysis(selected_analysis, is_parametric)
            analysis_window.destroy()

        ttk.Button(analysis_window, text="Виконати аналіз", command=run_selected_analysis).pack(pady=10)

    def perform_and_report_analysis(self, analysis_type, is_parametric):
        """Виконує обраний аналіз та генерує звіт."""
        try:
            df_cleaned = self.df.copy()
            numeric_df = df_cleaned.apply(pd.to_numeric, errors='coerce').select_dtypes(include=np.number)
            
            if analysis_type == "Дисперсійний аналіз":
                if numeric_df.shape[1] < 1 or df_cleaned.shape[1] < 2:
                    raise ValueError("Недостатньо даних для дисперсійного аналізу.")
                
                # Визначаємо залежну змінну (перший числовий стовпець) та фактор (перший текстовий)
                response_col = numeric_df.columns[0]
                factor_cols = df_cleaned.select_dtypes(exclude=np.number).columns
                if not factor_cols.empty:
                    factor_col = factor_cols[0]
                    formula = f'{response_col} ~ C({factor_col})' # C() позначає категоріальну змінну
                else:
                    raise ValueError("Відсутні текстові стовпці для використання як фактори.")

                model_df = pd.DataFrame({
                    response_col: numeric_df[response_col],
                    factor_col: df_cleaned[factor_col]
                }).dropna()

                model = ols(formula, data=model_df).fit()
                analysis_results = sm.stats.anova_lm(model, typ=2)
            
            elif analysis_type == "Регресія":
                if numeric_df.shape[1] < 2:
                    raise ValueError("Недостатньо числових стовпців для регресійного аналізу.")
                
                y = numeric_df.iloc[:, 0]
                X = sm.add_constant(numeric_df.iloc[:, 1])
                model = sm.OLS(y, X).fit()
                analysis_results = model.summary()
            
            else:
                analysis_results = f"Аналіз '{analysis_type}' ще не реалізовано."

            report = ReportGenerator(self.df, analysis_type, analysis_results)
            report.generate()
            
        except Exception as e:
            messagebox.showerror("Помилка", f"Не вдалося виконати аналіз: {e}")
