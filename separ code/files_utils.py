import pandas as pd
import os
import logging


# Проверка, что файл является табличным документом
def is_valid_file(file_path):
    valid_extensions = ['.xlsx', '.xls', '.csv']  # Разрешенные форматы файлов
    _, ext = os.path.splitext(file_path)
    return ext in valid_extensions


# Загрузка данных в зависимости от типа файла
def load_data(file_path, sheet_name=None):
    _, ext = os.path.splitext(file_path)

    if ext == '.xlsx' or ext == '.xls':
        return pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
    elif ext == '.csv':
        return pd.read_csv(file_path)
    else:
        raise ValueError(f"Формат файла {ext} не поддерживается.")


# Функция для создания уникальной папки
def create_unique_folder(base_folder='Graphics'):
    folder_name = base_folder
    version = 0

    # Проверка существования папки, если она уже есть, создается с номером версии
    while os.path.exists(folder_name):
        version += 1
        folder_name = f"{base_folder}{version}"

    # Создание новой папки
    os.makedirs(folder_name)
    return folder_name


# Функция для создания уникального лог-файла
def create_unique_log_file(base_log_file='errors.log'):
    log_file_name = base_log_file
    version = 0

    # Проверяем существование файла, если он уже есть, создаем с номером версии
    while os.path.exists(log_file_name):
        version += 1
        log_file_name = f"errors_{version}.log"

    return log_file_name

def load_valid_sheets(file_path):
    _, ext = os.path.splitext(file_path)

    if ext == '.xlsx' or ext == '.xls':
        xlsx = pd.ExcelFile(file_path)
        sheet_names = xlsx.sheet_names
        sheet_names = [name for name in sheet_names if 'Рис' in name]

        valid_sheets = []
        for sheet_name in sheet_names:
            try:
                df = load_data(file_path, sheet_name)

                # Проверяем, можно ли преобразовать необходимые столбцы в float
                df.iloc[2:, 2].astype(float)  # Данные за 2024 год (столбец C)
                df.iloc[2:, 3].astype(float)  # Данные за 2023 год (столбец D)
                df.iloc[2:, 4].astype(float)  # Данные за 2022 год (столбец E)

                valid_sheets.append(sheet_name)

            except Exception as e:
                # Логируем ошибку в файл
                logging.error(f"Ошибка при обработке листа '{sheet_name}': {str(e)}")
                print(f"Ошибка при обработке листа '{sheet_name}': {e}")
                continue

        return valid_sheets
    else:
        return [None]  # Для CSV просто используем имя файла без листов


def read_parameters_from_file(file_path):
    params = {}
    with open(file_path, 'r') as file:
        for line in file:
            key, value = line.strip().split('=')
            if key == "standard_deviation":
                params[key] = float(value)
            elif key == "show_original_values":
                params[key] = value.lower() == 'true'
    return params