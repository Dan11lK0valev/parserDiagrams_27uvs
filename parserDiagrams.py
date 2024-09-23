import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import ScalarFormatter
import matplotlib.ticker as ticker
import os
import logging
import sys


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


# Функция для создания уникльной папки
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


# Настройка логирования с уникальным именем файла
log_file_name = create_unique_log_file()
logging.basicConfig(filename=log_file_name, level=logging.ERROR,
                    format='%(asctime)s - %(levelname)s - %(message)s')


# основная функция реализации диаграмм
def make_diagrams(sheet_name, file_path, folder_name):
    df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')

    regions = df.iloc[2:, 0].values  # Региональные названия из столбца A, начиная с 4 строки
    values_2022 = df.iloc[2:, 4].astype(float).values  # Данные за 2022 год (столбец E)
    values_2023 = df.iloc[2:, 3].astype(float).values  # Данные за 2023 год (столбец D)
    values_2024 = df.iloc[2:, 2].astype(float).values  # Данные за 2024 год (столбец C)

    # Фильтруем регионы, у которых значения за все три года равны нулю
    non_zero_filter = (values_2022 != 0) | (values_2023 != 0) | (values_2024 != 0)
    regions = regions[non_zero_filter]
    values_2022 = values_2022[non_zero_filter]
    values_2023 = values_2023[non_zero_filter]
    values_2024 = values_2024[non_zero_filter]

    # Сортировка регионов по возрастанию значений за 2024 год
    sorted_indices = values_2024.argsort()
    regions = regions[sorted_indices]
    values_2022 = values_2022[sorted_indices]
    values_2023 = values_2023[sorted_indices]
    values_2024 = values_2024[sorted_indices]

    # Расчет средних значений для каждого года без учета нулевых значений
    mean_2024 = values_2024[values_2024 > 0].mean()
    mean_2023 = values_2023[values_2023 > 0].mean()
    mean_2022 = values_2022[values_2022 > 0].mean()

    # Настройка положения для каждой группы столбиков
    x = range(len(regions))
    width = 0.25  # Ширина каждого столбика

    # Построение диаграммы
    plt.figure(figsize=(10, 6), dpi=500)

    # Столбики для каждого года
    plt.bar([i - width for i in x], values_2022, width=width, label='2022', color='#A5A5A5')
    plt.bar(x, values_2023, width=width, label='2023', color='#ED7D31')
    plt.bar([i + width for i in x], values_2024, width=width, label='2024', color='#5B9BD5')

    # Горизонтальные пунктирные линии со средними значениями
    plt.axhline(y=mean_2022, color='#A5A5A5', linestyle='--', linewidth=0.8,
                label=f'Среднее за 2022: {int(round(mean_2022)):,}'.replace(',', ' '))
    plt.axhline(y=mean_2023, color='#ED7D31', linestyle='--', linewidth=0.8,
                label=f'Среднее за 2023: {int(round(mean_2023)):,}'.replace(',', ' '))
    plt.axhline(y=mean_2024, color='#5B9BD5', linestyle='--', linewidth=0.8,
                label=f'Среднее за 2024: {int(round(mean_2024)):,}'.replace(',', ' '))

    plt.xticks(x, regions, rotation=90, ha='right',
               fontsize=6)  # Поворот подписи на оси X на 90 градусов и настройка размера шрифта
    plt.rcParams['text.antialiased'] = True  # Сглаженность шрифтов
    plt.legend()
    plt.tight_layout()

    # Удаление лишних границ графика
    plt.gca().spines['right'].set_visible(False)
    plt.gca().spines['top'].set_visible(False)
    # Отключение научной нотации на оси Y
    plt.gca().yaxis.set_major_formatter(ScalarFormatter(useOffset=False))
    plt.ticklabel_format(style='plain', axis='y')  # Установка обычного стиля на оси Y

    # Форматирование оси ординат с пробелами в числах
    ax = plt.gca()
    ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, _: f'{int(x):,}'.replace(',', ' ')))

    plt.savefig(f'{folder_name}/{sheet_name}.png', dpi=500, bbox_inches='tight')
    plt.close()


# функция проверки листов табличных документов на валидность (читабельность)
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


# основная функция для начала работы программы
def main():
    # Загрузка табличного документа
    # Если есть аргументы командной строки
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        file_path = input(
            "Введите название табличного документа с его форматом\n(при его нахождение в одной директории с "
            "программой)\nили введите путь к нему вплоть до самого документа: ")

    if not is_valid_file(file_path):
        print("Ошибка: Выбранный файл не является табличным документом. Пожалуйста, выберите файл с расширением "
              ".xlsx, .xls или .csv.")
        sys.exit(1)

    # Создаем уникальную папку Graphics№
    folder_name = create_unique_folder()

    # Загружаем только валидные листы
    valid_sheets = load_valid_sheets(file_path)
    if valid_sheets:
        print("Валидные листы:", ', '.join(valid_sheets) if valid_sheets[0] else "CSV файл")

        for name in valid_sheets:
            make_diagrams(name, file_path, folder_name)
    else:
        print("Нет валидных листов для обработки.")

    print("Program has done. Thank you for using.")
    print("parserDiagram_27uvs.")


# Запуск программы
if __name__ == '__main__':
    main()
