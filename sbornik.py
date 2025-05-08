import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Tk, Label
from datetime import datetime
import logging
import sys
from ctypes import windll
from PIL import ImageTk, Image
import os

logging.basicConfig(level=logging.INFO)

windll.shcore.SetProcessDpiAwareness(1)

class ExcelProcessor:
    def __init__(self, file_path):
        self.file_path = file_path
        self.sheets = {}
        self.current_df = None
        self.load_sheets()

    def load_sheets(self):
        try:
            xls = pd.ExcelFile(self.file_path)
            self.sheets = {i: sheet for i, sheet in enumerate(xls.sheet_names, 1)}
        except Exception as e:
            logging.error(f"Ошибка загрузки файла: {e}")
            raise

    def load_sheet_data(self, sheet_number):
        try:
            sheet_name = self.sheets[sheet_number]
            self.current_df = pd.read_excel(self.file_path, sheet_name=sheet_name)
            return self.current_df.columns.tolist()
        except KeyError as e:
            logging.error(f"Неверный номер листа: {e}")
            raise

    @staticmethod
    def process_hours(x):
        try:
            if pd.isnull(x):
                return 0.0
            x = round(float(x), 2)
            rules = [
                (0, 0, 0),
                (0.01, 2.99, 0),
                (3.00, 3.99, 0.25),
                (4.00, 4.99, 0.5),
                (5.00, 6.99, 0.75),
                (7.00, 9.99, 1.0),
                (10.00, 10.99, 1.25),
                (11.00, 12.99, 1.5),
                (13.00, 14.99, 1.75),
                (15.00, 15.99, 2.0),
                (16.00, 17.99, 2.5),
                (18.00, 18.99, 2.75),
                (19.00, 20.99, 3.0),
                (21.00, 22.99, 3.25),
                (23.00, 24.99, 3.5),
            ]
            for min_val, max_val, deduction in rules:
                if min_val <= x <= max_val:
                    return round(x - deduction, 2)
            return 0.0
        except Exception as e:
            logging.error(f"Ошибка обработки часов: {e}")
            return 0.0

    def calculate_bonus(self, columns_mapping):
        try:
            df = self.current_df.rename(columns=columns_mapping)
            self._check_columns(df, ["Логин", "Сдельный заработок", "csat", "quality"])

            max_c = df["Сдельный заработок"].max()
            df["Вес сделки"] = df["Сдельный заработок"] / max_c * 0.1
            df["Вес CSAT"] = df["csat"] / 5 * 0.4
            df["Вес Quality"] = df["quality"] / 100 * 0.5

            df["Итоговый балл"] = df[["Вес сделки", "Вес CSAT", "Вес Quality"]].sum(
                axis=1
            )
            df_sorted = df.sort_values("Итоговый балл", ascending=False).reset_index(
                drop=True
            )

            n = len(df_sorted)
            boundaries = [0, 0.05, 0.15, 0.30, 0.50, 0.75, 1.0]
            positions = [1, 2, 3, 4, 5, 6]
            df_sorted["Ранг"] = 0

            for i in range(n):
                current_ratio = (i + 1) / n
                for pos in range(len(positions)):
                    if current_ratio <= boundaries[pos + 1]:
                        df_sorted.at[i, "Ранг"] = positions[pos]
                        break

            bonus_percent = {1: 35, 2: 25, 3: 15, 4: 10, 5: 5, 6: 0}
            df_sorted["Процент премии"] = df_sorted["Ранг"].map(bonus_percent)

            return df_sorted[
                [
                    "Ранг",
                    "Логин",
                    "Вес сделки",
                    "Вес CSAT",
                    "Вес Quality",
                    "Итоговый балл",
                    "Процент премии",
                ]
            ]

        except Exception as e:
            logging.error(f"Ошибка расчета премии: {e}")
            raise

    def calculate_hours(self, columns_mapping):
        try:
            df = self.current_df.rename(columns=columns_mapping)
            self._check_columns(
                df,
                [
                    "Логин",
                    "Теги",
                    "Тип",
                    "Начало (дата)",
                    "Начало (время)",
                    "Конец (дата)",
                    "Конец (время)",
                ],
            )

            df["Начало (дата)"] = pd.to_datetime(
                df["Начало (дата)"], format="%d.%m.%Y", dayfirst=True
            )
            df["Конец (дата)"] = pd.to_datetime(
                df["Конец (дата)"], format="%d.%m.%Y", dayfirst=True
            )
            df["Начало (время)"] = pd.to_datetime(
                df["Начало (время)"], format="%H:%M:%S"
            ).dt.time
            df["Конец (время)"] = pd.to_datetime(
                df["Конец (время)"], format="%H:%M:%S"
            ).dt.time

            # Создание меток времени
            df["Начало"] = df.apply(
                lambda row: pd.Timestamp.combine(
                    row["Начало (дата)"], row["Начало (время)"]
                ),
                axis=1,
            )
            df["Конец"] = df.apply(
                lambda row: pd.Timestamp.combine(
                    row["Конец (дата)"], row["Конец (время)"]
                ),
                axis=1,
            )

            shifts_mask = df["Тип"].isin(
                ["Смена. Основная", "Смена. Доп", "Смена. Отработка", "Сегмент смены"]
            )
            shifts_df = df[shifts_mask].copy()

            shifts_df["time_diff"] = shifts_df["Конец"] - shifts_df["Начало"]
            shifts_df["hours"] = shifts_df["time_diff"].dt.total_seconds() / 3600
            shifts_df["Часы с перерывами"] = shifts_df["hours"].apply(
                self.process_hours
            )

            violations_mask = (
                df["Тип"].str.startswith("Наставничество.")
                | (
                    df["Тип"].str.startswith("ПА.")
                    & ~df["Тип"].eq("ПА. Ошибочное нарушение")
                    & ~df["Тип"].str.startswith("Отсутствие.")
                )
            ) | df["Тип"].isin(
                [
                    "Нарушение. Не работает",
                    "Нарушение. Прогул",
                    "Нарушение. Опоздание на смену",
                ]
            )

            violations_df = df[violations_mask].copy()
            violations_df["time_diff"] = (
                violations_df["Конец"] - violations_df["Начало"]
            )
            violations_df["hours"] = (
                violations_df["time_diff"].dt.total_seconds() / 3600
            )
            violations_df["Часы с перерывами"] = violations_df["hours"].apply(
                self.process_hours
            )
            violations_df["Нарушения"] = violations_df["Часы с перерывами"] * -1

            # Группировка и объединение результатов
            group_cols = ["Логин", "Теги"]

            shifts_agg = (
                shifts_df.groupby(group_cols)["Часы с перерывами"].sum().reset_index()
            )
            violations_agg = (
                violations_df.groupby(group_cols)["Нарушения"].sum().reset_index()
            )

            merged = pd.merge(shifts_agg, violations_agg, on=group_cols, how="outer")
            merged.fillna(0, inplace=True)
            merged["Чистые часы"] = merged["Часы с перерывами"] + merged["Нарушения"]

            return merged

        except Exception as e:
            logging.error(f"Ошибка расчета часов: {e}")
            raise

    def _check_columns(self, df, required):
        missing = [col for col in required if col not in df.columns]
        if missing:
            raise ValueError(f"Отсутствуют столбцы: {', '.join(missing)}")

    def save_results(self, df, output_path):
        try:
            with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                df.to_excel(writer, index=False)
            logging.info(f"Файл сохранен: {output_path}")
        except Exception as e:
            logging.error(f"Ошибка сохранения: {e}")
            raise


class GUIApplication(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Аналитика 0.1")
        self.geometry("650x650")
        self.processor = None
        self.result_df = None
        self._init_ui()

    def _init_ui(self):
        self._create_notebook()
        self._create_status_bar()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _create_notebook(self):
        self.notebook = ttk.Notebook(self)
        self._init_file_tab()
        self._init_calc_tab()
        self.notebook.pack(expand=1, fill="both")

    def _init_file_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Файл")

        frame = ttk.LabelFrame(tab, text="Управление файлами")
        frame.pack(padx=10, pady=10, fill="both", expand=True)

        ttk.Button(frame, text="Выбрать файл", command=self._load_file).pack(pady=5)

        self.file_label = ttk.Label(frame, text="Файл не выбран")
        self.file_label.pack(pady=5)

        self.sheet_combo = ttk.Combobox(frame, state="readonly")
        self.sheet_combo.pack(pady=5)

        ttk.Button(frame, text="Загрузить данные", command=self._load_sheet).pack(
            pady=5
        )

    def _init_calc_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Расчеты")

        frame = ttk.LabelFrame(tab, text="Настройки расчетов")
        frame.pack(padx=10, pady=10, fill="both", expand=True)

        self.calc_type = tk.StringVar(value="bonus")
        ttk.Radiobutton(
            frame,
            text="Премия",
            variable=self.calc_type,
            value="bonus",
            command=self._update_columns,
        ).grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(
            frame,
            text="Часы",
            variable=self.calc_type,
            value="hours",
            command=self._update_columns,
        ).grid(row=0, column=1, sticky="w")

        self.columns_frame = ttk.Frame(frame)
        self.columns_frame.grid(row=1, column=0, columnspan=2, pady=10, sticky="nsew")

        ttk.Button(frame, text="Рассчитать", command=self._run_calculation).grid(
            row=2, column=0, pady=10
        )
        ttk.Button(frame, text="Сохранить", command=self._save_results).grid(
            row=2, column=1, pady=10
        )

        self._update_columns()

    def _create_status_bar(self):
        self.status = tk.StringVar()
        ttk.Label(self, textvariable=self.status, relief="sunken", anchor="w").pack(
            side="bottom", fill="x"
        )

    def _on_close(self):
        self.destroy()

    def _load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if file_path:
            try:
                self.processor = ExcelProcessor(file_path)
                self.file_label.config(text=file_path)
                self.sheet_combo["values"] = list(self.processor.sheets.values())
                self.status.set("Файл загружен")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка загрузки: {str(e)}")

    def _load_sheet(self):
        try:
            sheet_name = self.sheet_combo.get()
            if not sheet_name:
                raise ValueError("Выберите лист")

            sheet_num = [
                k for k, v in self.processor.sheets.items() if v == sheet_name
            ][0]
            columns = self.processor.load_sheet_data(sheet_num)
            self._update_columns()
            self.status.set(f"Загружено столбцов: {len(columns)}")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def _update_columns(self):
        try:
            for widget in self.columns_frame.winfo_children():
                widget.destroy()

            columns_config = {
                "bonus": [
                    {"name": "Логин", "required": True},
                    {"name": "Сдельный заработок", "required": True},
                    {"name": "CSAT", "required": True},
                    {"name": "Quality", "required": True},
                ],
                "hours": [
                    {"name": "Логин", "required": True},
                    {"name": "Теги", "required": True},
                    {"name": "Тип", "required": True},
                    {"name": "Начало (дата)", "required": True},
                    {"name": "Конец (дата)", "required": True},
                    {"name": "Начало (время)", "required": True},
                    {"name": "Конец (время)", "required": True},
                ],
            }

            calc_type = self.calc_type.get()
            columns = (
                self.processor.current_df.columns.tolist() if self.processor else []
            )

            self.column_combos = {}
            for row, field_config in enumerate(columns_config[calc_type]):
                frame = ttk.Frame(self.columns_frame)
                frame.grid(row=row, column=0, sticky="ew", pady=2)

                label_text = f"{field_config['name']}:"
                if field_config["required"]:
                    label_text += " *"

                ttk.Label(frame, text=label_text).pack(side="left", padx=5)

                combo = ttk.Combobox(frame, values=columns, state="readonly")
                combo.pack(side="right", fill="x", expand=True)

                if field_config["name"] in columns:
                    combo.set(field_config["name"])

                self.column_combos[field_config["name"]] = combo
        except Exception as e:
            logging.error(f"Ошибка обновления: {e}")

    def _run_calculation(self):
        try:
            calc_type = self.calc_type.get()
            mapping = {k: v.get() for k, v in self.column_combos.items()}

            required_fields = [
                field["name"]
                for field in self._get_current_config()
                if field["required"]
            ]

            missing = [field for field in required_fields if not mapping.get(field)]
            if missing:
                raise ValueError(
                    f"Обязательные поля не заполнены:\n{', '.join(missing)}"
                )

            if calc_type == "bonus":
                self.result_df = self.processor.calculate_bonus(mapping)
            else:
                self.result_df = self.processor.calculate_hours(mapping)

            messagebox.showinfo("Успех", "Расчет завершен")
            self.status.set("Результаты готовы")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка расчета:\n{str(e)}")
            logging.error(str(e))

    def _get_current_config(self):
        calc_type = self.calc_type.get()
        return {
            "bonus": [
                {"name": "Логин", "required": True},
                {"name": "Сдельный заработок", "required": True},
                {"name": "CSAT", "required": True},
                {"name": "Quality", "required": True},
            ],
            "hours": [
                {"name": "Логин", "required": True},
                {"name": "Теги", "required": True},
                {"name": "Тип", "required": True},
                {"name": "Начало (дата)", "required": True},
                {"name": "Конец (дата)", "required": True},
                {"name": "Начало (время)", "required": True},
                {"name": "Конец (время)", "required": True},
            ],
        }[calc_type]
    
    def resource_path(relative_path):
        try:
                base_path = sys._MEIPASS
        except Exception:
                base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

        # Иконка
    icon_path = resource_path("icon.ico")

    def _save_results(self):
        if self.result_df is not None:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx"), ("All files", "*.*")],
            )
            if file_path:
                try:
                    self.processor.save_results(self.result_df, file_path)
                    messagebox.showinfo("Успех", "Файл сохранен")
                    self.status.set("Сохранено: " + file_path)
                except Exception as e:
                    messagebox.showerror("Ошибка", str(e))
        else:
            messagebox.showwarning("Внимание", "Сначала выполните расчет")


if __name__ == "__main__":
    app = GUIApplication()
    app.mainloop()
