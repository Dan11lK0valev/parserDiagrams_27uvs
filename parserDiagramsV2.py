import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import ScalarFormatter
import matplotlib.ticker as ticker
import os
import logging
import sys


# Набор функций, отвечающие за чтение и анализ данных из файлов
class FileUtils:
    # Функция проверки файл на то, что он табличного типа
    @staticmethod
    def is_valid_file(file_path):
        valid_extensions = ['.xlsx', '.xls', '.csv']
        _, ext = os.path.splitext(file_path)
        return ext in valid_extensions

    # Функция загрузки данных из табличного документа
    @staticmethod
    def load_data(file_path, sheet_name=None):
        _, ext = os.path.splitext(file_path)
        if ext == '.xlsx' or ext == '.xls':
            return pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
        elif ext == '.csv':
            return pd.read_csv(file_path)
        else:
            raise ValueError(f"Формат файла {ext} не поддерживается.")

    # Функция создания папки для сохранения png картинок сгенерированных диаграмм
    @staticmethod
    def create_unique_folder(base_folder='Graphics'):
        folder_name = base_folder
        version = 0
        while os.path.exists(folder_name):
            version += 1
            folder_name = f"{base_folder}{version}"
        os.makedirs(folder_name)
        return folder_name

    # Функция отвечающая за создание error.log файлов
    @staticmethod
    def create_unique_log_file(base_log_file='errors.log'):
        log_file_name = base_log_file
        version = 0
        while os.path.exists(log_file_name):
            version += 1
            log_file_name = f"errors_{version}.log"
        return log_file_name


# Набор функций, отвечающие за сбор данных
class DataProcessor:
    # Сбор читаемых листов из табличного документа
    @staticmethod
    def load_valid_sheets(file_path):
        _, ext = os.path.splitext(file_path)
        if ext == '.xlsx' or ext == '.xls':
            xlsx = pd.ExcelFile(file_path)
            sheet_names = [name for name in xlsx.sheet_names if 'Рис' in name]
            valid_sheets = []

            for sheet_name in sheet_names:
                try:
                    df = FileUtils.load_data(file_path, sheet_name)
                    df.iloc[2:, 2].astype(float)
                    df.iloc[2:, 3].astype(float)
                    df.iloc[2:, 4].astype(float)
                    valid_sheets.append(sheet_name)
                except Exception as e:
                    logging.error(f"Ошибка при обработке листа '{sheet_name}': {str(e)}")
                    print(f"Ошибка при обработке листа '{sheet_name}': {e}")
                    continue
            return valid_sheets
        else:
            # В случае CSV файла нет листов
            return [None]

    # Чтение двух параметров из файла parameters.txt
    @staticmethod
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


# Набор функций отвечающих за генерацию диаграмм
class DiagramConstructor:
    # Нахождение значений превышающих стандартное отклонение и их сокращение
    @staticmethod
    def bar_adjust(value, threshold, max_non_outlier_value):
        adjusted_height = value
        white_cut = False
        if value > threshold:
            if value > max_non_outlier_value:
                adjusted_height = max_non_outlier_value
                white_cut = True
        return adjusted_height, white_cut

    # Добавление "белых срезов" для сокращенных столбцов
    @staticmethod
    def add_white_section(bars, original_values, adjusted_values):
        for bar, original, (adjusted, white_cut) in zip(bars, original_values, adjusted_values):
            if white_cut:
                white_height = bar.get_height() * 0.02
                x = bar.get_x() + bar.get_width() / 2
                y = bar.get_height() - bar.get_height() * 0.05
                plt.bar(x, white_height, bottom=y, width=bar.get_width(), color='white', edgecolor='none')

    # Базовая функция генерации диаграмм
    @staticmethod
    def make_diagrams(sheet_name, file_path, folder_name, standard_deviation, show_original_values):
        df = FileUtils.load_data(file_path, sheet_name=sheet_name)

        regions = df.iloc[2:, 0].values  # Чтение регионов из колонки A с 4 строки
        values_1 = df.iloc[2:, 4].astype(float).values  # Чтение числовых значений из колонки E с 4 строки (это 2022)
        values_2 = df.iloc[2:, 3].astype(float).values  # Чтение числовых значений из колонки D с 4 строки (это 2023)
        values_3 = df.iloc[2:, 2].astype(float).values  # Чтение числовых значений из колонки C с 4 строки (это 2024)

        # Года из третьей строки (индекс 2) и соответствующих столбцов
        year_1 = df.iloc[1, 4]  # Год в колонке E (2022)
        year_2 = df.iloc[1, 3]  # Год в колонке D (2023)
        year_3 = df.iloc[1, 2]  # Год в колонке C (2024)

        non_zero_and_one_filter = ((values_1 != 0).astype(int) + (values_2 != 0).astype(int) +
                                   (values_3 != 0).astype(int)) > 1  # исключаем регионы без данных и с данными
        # за один год
        regions, values_1, values_2, values_3 = [arr[non_zero_and_one_filter] for arr in
                                                 [regions, values_1, values_2, values_3]]

        sorted_indices = values_3.argsort()  # сортировка регионов по возрастанию относительно колонки C
        regions, values_1, values_2, values_3 = [arr[sorted_indices] for arr in
                                                 [regions, values_1, values_2, values_3]]

        valid_data_filter = ((values_1 != 0).astype(int) + (values_2 != 0).astype(int) + (
                values_3 != 0).astype(int)) > 1
        mean_1, mean_2, mean_3 = [arr[valid_data_filter & (arr > 0)].mean() for arr in
                                  [values_1, values_2, values_3]]  # Средние значения по фильтрации
        std_2022, std_2023, std_2024 = [arr[valid_data_filter & (arr > 0)].std(ddof=0) for arr in
                                        [values_1, values_2, values_3]]  # Показатели стандартного отклонения
        # по фильтрации
        threshold_1, threshold_2, threshold_3 = [mean + standard_deviation * std for mean, std in
                                                 zip([mean_1, mean_2, mean_3],
                                                     [std_2022, std_2023, std_2024])]  # Пороги выбросов по фильтрации

        x = range(len(regions))
        width = 0.25

        plt.figure(figsize=(10, 6), dpi=500)
        # Максимальное значение, не преодолевшее порог выброса
        max_non_outlier_value = max(
            max(values_1[values_1 <= threshold_1]),
            max(values_2[values_2 <= threshold_2]),
            max(values_3[values_3 <= threshold_3])
        )
        # Генерация диаграммы
        adjusted_1 = [DiagramConstructor.bar_adjust(v, threshold_1, max_non_outlier_value) for v in values_1]
        bars_1 = plt.bar([i - width for i in x], [adj[0] for adj in adjusted_1], width=width,
                         label=f'{str(int(year_1))} — {int(round(mean_1)):,}'.replace(',', ' '), color='#A5A5A5')

        adjusted_2 = [DiagramConstructor.bar_adjust(v, threshold_2, max_non_outlier_value) for v in values_2]
        bars_2 = plt.bar(x, [adj[0] for adj in adjusted_2], width=width,
                         label=f'{str(int(year_2))} — {int(round(mean_2)):,}'.replace(',', ' '), color='#ED7D31')

        adjusted_3 = [DiagramConstructor.bar_adjust(v, threshold_3, max_non_outlier_value) for v in values_3]
        bars_3 = plt.bar([i + width for i in x], [adj[0] for adj in adjusted_3], width=width,
                         label=f'{str(int(year_3))} — {int(round(mean_3)):,}'.replace(',', ' '), color='#5B9BD5')

        # Подрисовка белых обрезаний у сокращенных столбиков
        DiagramConstructor.add_white_section(bars_1, values_1, adjusted_1)
        DiagramConstructor.add_white_section(bars_2, values_2, adjusted_2)
        DiagramConstructor.add_white_section(bars_3, values_3, adjusted_3)

        ax = plt.gca()
        y_ticks = ax.get_yticks()
        y_max_metric = 0
        if len(y_ticks) >= 2:
            y_max_metric = int(y_ticks[-2])  # Максимальная метрика из оси ординаты, используемая на графике

        # Визуализация числовых значений сокращенных и превышающих максимальную метрику оси Y столбцов
        if show_original_values:
            for i in range(len(x)):
                if ((values_2[i] > threshold_2 and values_2[i] > max_non_outlier_value) or
                        (values_2[i] > y_max_metric)):
                    plt.text(x[i],
                             DiagramConstructor.bar_adjust(values_2[i], threshold_2, max_non_outlier_value)[0] +
                             3.5, f'{int(values_2[i]):,}'.replace(',', ' '), ha='center', va='bottom',
                             rotation=90, fontsize=5)
                if ((values_1[i] > threshold_1 and values_1[i] > max_non_outlier_value) or
                        (values_1[i] > y_max_metric)):
                    plt.text(x[i] - width,
                             DiagramConstructor.bar_adjust(values_1[i], threshold_1, max_non_outlier_value)[0] +
                             3.5, f'{int(values_1[i]):,}'.replace(',', ' '), ha='right', va='bottom',
                             rotation=90, fontsize=5)
                if ((values_3[i] > threshold_3 and values_3[i] > max_non_outlier_value) or
                        (values_3[i] > y_max_metric)):
                    plt.text(x[i] + width,
                             DiagramConstructor.bar_adjust(values_3[i], threshold_3, max_non_outlier_value)[0] +
                             3.5, f'{int(values_3[i]):,}'.replace(',', ' '), ha='left', va='bottom',
                             rotation=90, fontsize=5)

        # Генерация пунктирных линий - средних значений за каждый год
        plt.axhline(y=mean_1, color='#A5A5A5', linestyle='--', linewidth=0.8)
        plt.axhline(y=mean_2, color='#ED7D31', linestyle='--', linewidth=0.8)
        plt.axhline(y=mean_3, color='#5B9BD5', linestyle='--', linewidth=0.8)

        # Настройки для повышенного качества, нормализации названия регионов на оси X и числового формата метрик оси Y
        plt.xticks(x, regions, rotation=90, ha='right', fontsize=6)
        plt.rcParams['text.antialiased'] = True
        plt.legend(title='Среднее', labelcolor=['#A5A5A5', '#ED7D31', '#5B9BD5'], handlelength=0, frameon=False)
        plt.tight_layout()
        plt.gca().spines['right'].set_visible(False)
        plt.gca().spines['top'].set_visible(False)
        plt.gca().yaxis.set_major_formatter(ScalarFormatter(useOffset=False))
        plt.ticklabel_format(style='plain', axis='y')
        ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, _: f'{int(x):,}'.replace(',', ' ')))

        # Сохранение диаграммы
        plt.savefig(f'{folder_name}/{sheet_name}.png', dpi=500, bbox_inches='tight')
        plt.close()


# Главная база программы
class MainApp:
    # Проверка читаемого объекта на существование и верный формат
    @staticmethod
    def run():
        if len(sys.argv) > 1:
            file_path = sys.argv[1]
        else:
            file_path = input(
                "Введите название табличного документа с его форматом\n(при его нахождение в одной директории с "
                "программой)\nили введите путь к нему вплоть до самого документа: ")

        if not FileUtils.is_valid_file(file_path):
            print("Ошибка: Неверный формат файла.")
            sys.exit(1)

        folder_name = FileUtils.create_unique_folder()
        valid_sheets = DataProcessor.load_valid_sheets(file_path)
        # Анализ листов табличного документа, которые код прочитывает корректно
        if valid_sheets:
            print("Валидные листы:", ', '.join(valid_sheets) if valid_sheets[0] else "CSV файл")
            # Параметры генерации диаграмм из файла parameters.txt
            params = DataProcessor.read_parameters_from_file('parameters.txt')
            standard_deviation = params.get('standard_deviation', 4)
            show_original_values = params.get('show_original_values', True)
            # Запуск генерации диаграмм
            for sheet in valid_sheets:
                DiagramConstructor.make_diagrams(sheet, file_path, folder_name, standard_deviation,
                                                 show_original_values)
        else:
            print("Нет валидных листов для обработки.")
        # Завершение программы
        print("Program has done. Thank you for using.")
        print("parserDiagram_27uvs.")


# Запуск программы
if __name__ == '__main__':
    log_file_name = FileUtils.create_unique_log_file()
    logging.basicConfig(filename=log_file_name, level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')
    MainApp.run()
