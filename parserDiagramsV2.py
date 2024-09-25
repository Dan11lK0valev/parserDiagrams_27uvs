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
        values_2022 = df.iloc[2:, 4].astype(float).values  # Чтение числовых значений за 2022 из колонки E с 4 строки
        values_2023 = df.iloc[2:, 3].astype(float).values  # Чтение числовых значений за 2023 из колонки D с 4 строки
        values_2024 = df.iloc[2:, 2].astype(float).values  # Чтение числовых значений за 2024 из колонки C с 4 строки

        non_zero_and_one_filter = ((values_2022 != 0).astype(int) + (values_2023 != 0).astype(int) +
                                   (values_2024 != 0).astype(int)) > 1  # исключаем регионы без данных и с данными
        # за один год
        regions, values_2022, values_2023, values_2024 = [arr[non_zero_and_one_filter] for arr in
                                                          [regions, values_2022, values_2023, values_2024]]

        sorted_indices = values_2024.argsort()  # сортировка регионов по возрастанию относительно колонки C
        regions, values_2022, values_2023, values_2024 = [arr[sorted_indices] for arr in
                                                          [regions, values_2022, values_2023, values_2024]]

        valid_data_filter = ((values_2022 != 0).astype(int) + (values_2023 != 0).astype(int) + (
                values_2024 != 0).astype(int)) > 1
        mean_2022, mean_2023, mean_2024 = [arr[valid_data_filter & (arr > 0)].mean() for arr in
                                           [values_2022, values_2023, values_2024]]
        std_2022, std_2023, std_2024 = [arr[valid_data_filter & (arr > 0)].std(ddof=0) for arr in
                                        [values_2022, values_2023, values_2024]]
        threshold_2022, threshold_2023, threshold_2024 = [mean + standard_deviation * std for mean, std in
                                                          zip([mean_2022, mean_2023, mean_2024],
                                                              [std_2022, std_2023, std_2024])]

        x = range(len(regions))
        width = 0.25

        plt.figure(figsize=(10, 6), dpi=500)
        max_non_outlier_value = max(
            max(values_2022[values_2022 <= threshold_2022]),
            max(values_2023[values_2023 <= threshold_2023]),
            max(values_2024[values_2024 <= threshold_2024])
        )

        adjusted_2022 = [DiagramConstructor.bar_adjust(v, threshold_2022, max_non_outlier_value) for v in values_2022]
        bars_2022 = plt.bar([i - width for i in x], [adj[0] for adj in adjusted_2022], width=width, label='2022',
                            color='#A5A5A5')

        adjusted_2023 = [DiagramConstructor.bar_adjust(v, threshold_2023, max_non_outlier_value) for v in values_2023]
        bars_2023 = plt.bar(x, [adj[0] for adj in adjusted_2023], width=width, label='2023', color='#ED7D31')

        adjusted_2024 = [DiagramConstructor.bar_adjust(v, threshold_2024, max_non_outlier_value) for v in values_2024]
        bars_2024 = plt.bar([i + width for i in x], [adj[0] for adj in adjusted_2024], width=width, label='2024',
                            color='#5B9BD5')

        DiagramConstructor.add_white_section(bars_2022, values_2022, adjusted_2022)
        DiagramConstructor.add_white_section(bars_2023, values_2023, adjusted_2023)
        DiagramConstructor.add_white_section(bars_2024, values_2024, adjusted_2024)

        ax = plt.gca()
        y_ticks = ax.get_yticks()
        y_max_metric = 0
        if len(y_ticks) >= 2:
            y_max_metric = int(y_ticks[-2])

        if show_original_values:
            for i in range(len(x)):
                if ((values_2023[i] > threshold_2023 and values_2023[i] > max_non_outlier_value) or
                        (values_2023[i] > y_max_metric)):
                    plt.text(x[i],
                             DiagramConstructor.bar_adjust(values_2023[i], threshold_2023, max_non_outlier_value)[0] +
                             3.5, f'{int(values_2023[i]):,}'.replace(',', ' '), ha='center', va='bottom',
                             rotation=90, fontsize=5)
                if ((values_2022[i] > threshold_2022 and values_2022[i] > max_non_outlier_value) or
                        (values_2022[i] > y_max_metric)):
                    plt.text(x[i] - width,
                             DiagramConstructor.bar_adjust(values_2022[i], threshold_2022, max_non_outlier_value)[0] +
                             3.5, f'{int(values_2022[i]):,}'.replace(',', ' '), ha='right', va='bottom',
                             rotation=90, fontsize=5)
                if ((values_2024[i] > threshold_2024 and values_2024[i] > max_non_outlier_value) or
                        (values_2024[i] > y_max_metric)):
                    plt.text(x[i] + width,
                             DiagramConstructor.bar_adjust(values_2024[i], threshold_2024, max_non_outlier_value)[0] +
                             3.5, f'{int(values_2024[i]):,}'.replace(',', ' '), ha='left', va='bottom',
                             rotation=90, fontsize=5)

        plt.axhline(y=mean_2022, color='#A5A5A5', linestyle='--', linewidth=0.8,
                    label=f'Среднее за 2022: {int(round(mean_2022)):,}'.replace(',', ' '))
        plt.axhline(y=mean_2023, color='#ED7D31', linestyle='--', linewidth=0.8,
                    label=f'Среднее за 2023: {int(round(mean_2023)):,}'.replace(',', ' '))
        plt.axhline(y=mean_2024, color='#5B9BD5', linestyle='--', linewidth=0.8,
                    label=f'Среднее за 2024: {int(round(mean_2024)):,}'.replace(',', ' '))

        plt.xticks(x, regions, rotation=90, ha='right', fontsize=6)
        plt.rcParams['text.antialiased'] = True
        plt.legend()
        plt.tight_layout()
        plt.gca().spines['right'].set_visible(False)
        plt.gca().spines['top'].set_visible(False)
        plt.gca().yaxis.set_major_formatter(ScalarFormatter(useOffset=False))
        plt.ticklabel_format(style='plain', axis='y')
        ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, _: f'{int(x):,}'.replace(',', ' ')))

        plt.savefig(f'{folder_name}/{sheet_name}.png', dpi=500, bbox_inches='tight')
        plt.close()


# Главная база программы
class MainApp:
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

        if valid_sheets:
            print("Валидные листы:", ', '.join(valid_sheets) if valid_sheets[0] else "CSV файл")
            params = DataProcessor.read_parameters_from_file('parameters.txt')
            standard_deviation = params.get('standard_deviation', 4)
            show_original_values = params.get('show_original_values', True)

            for sheet in valid_sheets:
                DiagramConstructor.make_diagrams(sheet, file_path, folder_name, standard_deviation,
                                                 show_original_values)
        else:
            print("Нет валидных листов для обработки.")

        print("Program has done. Thank you for using.")
        print("parserDiagram_27uvs.")


# Запуск программы
if __name__ == '__main__':
    log_file_name = FileUtils.create_unique_log_file()
    logging.basicConfig(filename=log_file_name, level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')
    MainApp.run()
