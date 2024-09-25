import matplotlib.pyplot as plt


# Функция для определения возможного сокращения
def bar_adjust(value, threshold, max_non_outlier_value):
    adjusted_height = value  # Значение по умолчанию для высоты столбца
    white_cut = False

    if value > threshold:
        # Сокращаем высоту столбца до максимального невыбросного значения
        if value > max_non_outlier_value:
            adjusted_height = max_non_outlier_value
            white_cut = True

    return adjusted_height, white_cut


# Функция для добавления белых частей в сокращенные столбики
def add_white_section(bars, original_values, adjusted_values):
    for bar, original, (adjusted, white_cut) in zip(bars, original_values, adjusted_values):
        # Проверяем, был ли столбик сокращен
        if white_cut:
            # Рассчитываем высоту белой части
            white_height = bar.get_height() * 0.02
            # Получаем координаты текущего столбика
            x = bar.get_x() + bar.get_width() / 2
            y = bar.get_height() - bar.get_height() * 0.05

            # Добавляем белую часть (прямоугольник) только для сокращенных столбиков
            plt.bar(x, white_height, bottom=y, width=bar.get_width(), color='white', edgecolor='none')

