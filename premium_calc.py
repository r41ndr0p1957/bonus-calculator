import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from ctypes import windll
import sys
import os

windll.shcore.SetProcessDpiAwareness(1)

def load_file(): # Загружаем файл тут
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        entry.delete(0, tk.END)
        entry.insert(0, file_path)

def process_file(): # Обрабатываем и сохраняем тут
    file_path = entry.get()
    if not file_path:
        messagebox.showerror("Ошибка", "Выберите файл")
        return

    try: # Читаем лист "Премия" тут
        df = pd.read_excel(file_path, sheet_name='Премия')
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось прочитать файл: {str(e)}")
        return

    # Считаем
    max_c = df['Сдельный заработок'].max()
    df['Вес сделки'] = df['Сдельный заработок'] / max_c * 0.1
    df['Вес CSAT'] = df['csat'] / 5 * 0.4
    df['Вес Quality'] = df['quality'] / 100 * 0.5
    df['Итоговый балл'] = df[['Вес сделки', 'Вес CSAT', 'Вес Quality']].sum(axis=1)

    # Сортируем по убыванию итогового балла
    df_sorted = df.sort_values('Итоговый балл', ascending=False).reset_index(drop=True)
    n = len(df_sorted)

    # Расчет ранга саппорта
    boundaries = [0, 0.05, 0.15, 0.30, 0.50, 0.75, 1.0]
    positions = [1, 2, 3, 4, 5, 6]
    df_sorted['Ранг'] = 0

    for i in range(n):
        current_ratio = (i + 1) / n
        for pos in range(len(positions)):
            if current_ratio <= boundaries[pos + 1]:
                df_sorted.at[i, 'Ранг'] = positions[pos]
                break

     # Добавляем столбец с процентом премии
    bonus_percent = {
        1: 35,
        2: 25,
        3: 15,
        4: 10,
        5: 5,
        6: 0
    }
    df_sorted['Процент премии'] = df_sorted['Ранг'].map(bonus_percent)

    # Создаем итоговый DataFrame
    result_df = df_sorted[['Ранг', 'Логин', 'Вес сделки', 'Вес CSAT', 'Вес Quality', 'Итоговый балл', 'Процент премии']].copy()

    # Сохраняемся
    new_file_path = file_path.replace('.xlsx', '_processed.xlsx')
    try:
        with pd.ExcelWriter(new_file_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Премия', index=False)
            result_df.to_excel(writer, sheet_name='Результат', index=False)
        messagebox.showinfo("Готово", f"Файл сохранен как {new_file_path}")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка при сохранении: {str(e)}")

# Визуальная оболочка
root = tk.Tk()
root.title("Premium calc. V 0.1")

frame = tk.Frame(root)
frame.pack(padx=20, pady=20)

entry = tk.Entry(frame, width=50)
entry.pack(side=tk.TOP)

btn_load = tk.Button(frame, text="Загрузить файл", command=load_file)
btn_load.pack(pady=10)

btn_process = tk.Button(frame, text="Обработать файл", command=process_file)
btn_process.pack(pady=10)

def resource_path(relative_path):
    try:
                base_path = sys._MEIPASS
    except Exception:
                base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

        # Иконка
icon_path = resource_path("icon.ico")

root.mainloop()