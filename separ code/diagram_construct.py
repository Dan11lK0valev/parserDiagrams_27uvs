import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import ScalarFormatter
import matplotlib.ticker as ticker
from internal_functions import bar_adjust, add_white_section


# основная функция реализации диаграмм
def make_diagrams(sheet_name, file_path, folder_name, standard_deviation, show_original_values):
    df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')

    regions = df.iloc[2:, 0].values  # Региональные названия из столбца A, начиная с 4 строки
    values_2022 = df.iloc[2:, 4].astype(float).values  # Данные за 2022 год (столбец E)
    values_2023 = df.iloc[2:, 3].astype(float).values  # Данные за 2023 год (столбец D)
    values_2024 = df.iloc[2:, 2].astype(float).values  # Данные за 2024 год (столбец C)

    # Фильтруем регионы, у которых данные за все три года равны нулю или данные есть только за один год
    non_zero_and_one_filter = ((values_2022 != 0).astype(int) + (values_2023 != 0).astype(int) + (
                values_2024 != 0).astype(
        int)) > 1
    regions = regions[non_zero_and_one_filter]
    values_2022 = values_2022[non_zero_and_one_filter]
    values_2023 = values_2023[non_zero_and_one_filter]
    values_2024 = values_2024[non_zero_and_one_filter]

    # Сортировка регионов по возрастанию значений за 2024 год
    sorted_indices = values_2024.argsort()
    regions = regions[sorted_indices]
    values_2022 = values_2022[sorted_indices]
    values_2023 = values_2023[sorted_indices]
    values_2024 = values_2024[sorted_indices]

    # Фильтр для регионов, у которых есть данные хотя бы за два года
    valid_data_filter = ((values_2022 != 0).astype(int) + (values_2023 != 0).astype(int) + (values_2024 != 0).astype(
        int)) > 1
    # Расчет средних значений для каждого года с учетом только валидных данных
    mean_2022 = values_2022[valid_data_filter & (values_2022 > 0)].mean()
    mean_2023 = values_2023[valid_data_filter & (values_2023 > 0)].mean()
    mean_2024 = values_2024[valid_data_filter & (values_2024 > 0)].mean()

    std_2022 = values_2022[valid_data_filter & (values_2022 > 0)].std(ddof=0)
    std_2023 = values_2023[valid_data_filter & (values_2023 > 0)].std(ddof=0)
    std_2024 = values_2024[valid_data_filter & (values_2024 > 0)].std(ddof=0)

    threshold_2022 = mean_2022 + standard_deviation * std_2022
    threshold_2023 = mean_2023 + standard_deviation * std_2023
    threshold_2024 = mean_2024 + standard_deviation * std_2024

    # Настройка положения для каждой группы столбиков
    x = range(len(regions))
    width = 0.25  # Ширина каждого столбика

    # Построение диаграммы
    plt.figure(figsize=(10, 6), dpi=500)
    # fig, ax = plt.subplots()  # Создаем фигуру и оси

    # После вычисления порогов для выбросов
    max_non_outlier_value = max(
        max(values_2022[values_2022 <= threshold_2022]),
        max(values_2023[values_2023 <= threshold_2023]),
        max(values_2024[values_2024 <= threshold_2024])
    )

    # Столбики для каждого года
    adjusted_2022 = [bar_adjust(v, threshold_2022, max_non_outlier_value) for v in values_2022]
    bars_2022 = plt.bar([i - width for i in x], [adj[0] for adj in adjusted_2022], width=width, label='2022',
                        color='#A5A5A5')

    adjusted_2023 = [bar_adjust(v, threshold_2023, max_non_outlier_value) for v in values_2023]
    bars_2023 = plt.bar(x, [adj[0] for adj in adjusted_2023], width=width, label='2023', color='#ED7D31')

    adjusted_2024 = [bar_adjust(v, threshold_2024, max_non_outlier_value) for v in values_2024]
    bars_2024 = plt.bar([i + width for i in x], [adj[0] for adj in adjusted_2024], width=width, label='2024',
                        color='#5B9BD5')

    # Добавляем белые части только для сокращенных столбиков
    add_white_section(bars_2022, values_2022, adjusted_2022)
    add_white_section(bars_2023, values_2023, adjusted_2023)
    add_white_section(bars_2024, values_2024, adjusted_2024)

    ax = plt.gca()
    y_ticks = ax.get_yticks()
    y_max_metric = 0
    # Получение предпоследнего значения и преобразование в int
    if len(y_ticks) >= 2:  # Проверка, чтобы избежать ошибок
        y_max_metric = int(y_ticks[-2])

    # Если show_original_values = True, отображаем исходные значения над сокращенными и очень высокими столбиками
    if show_original_values:
        for i in range(len(x)):
            position_2022 = 'center'
            position_2023 = 'center'
            position_2024 = 'center'
            # Для 2023 года
            if (values_2023[i] > threshold_2023 and values_2023[i] > max_non_outlier_value) or (values_2023[i] > y_max_metric):
                position_2022 = "right"
                position_2024 = "left"
                plt.text(x[i],
                         bar_adjust(values_2023[i], threshold_2023, max_non_outlier_value)[0] + 3.5,
                         f'{int(values_2023[i]):,}'.replace(',', ' '),
                         ha=position_2023, va='bottom', rotation=90, fontsize=5)

            # Для 2022 года
            if (values_2022[i] > threshold_2022 and values_2022[i] > max_non_outlier_value) or (values_2022[i] > y_max_metric):
                plt.text(x[i] - width,
                         bar_adjust(values_2022[i], threshold_2022, max_non_outlier_value)[0] + 3.5,
                         f'{int(values_2022[i]):,}'.replace(',', ' '),
                         ha=position_2022, va='bottom', rotation=90, fontsize=5)

            # Для 2024 года
            if (values_2024[i] > threshold_2024 and values_2024[i] > max_non_outlier_value) or (values_2024[i] > y_max_metric):
                plt.text(x[i] + width,
                         bar_adjust(values_2024[i], threshold_2024, max_non_outlier_value)[0] + 3.5,
                         f'{int(values_2024[i]):,}'.replace(',', ' '),
                         ha=position_2024, va='bottom', rotation=90, fontsize=5)

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
    ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, _: f'{int(x):,}'.replace(',', ' ')))

    plt.savefig(f'{folder_name}/{sheet_name}.png', dpi=500, bbox_inches='tight')
    plt.close()