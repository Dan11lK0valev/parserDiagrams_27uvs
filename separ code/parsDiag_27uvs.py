from files_utils import (is_valid_file, create_unique_folder,
                         create_unique_log_file, load_valid_sheets, read_parameters_from_file)
from diagram_construct import make_diagrams
import sys
import logging


# Настройка логирования с уникальным именем файла
log_file_name = create_unique_log_file()
logging.basicConfig(filename=log_file_name, level=logging.ERROR,
                    format='%(asctime)s - %(levelname)s - %(message)s')


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

    # Создаем уникальную папку Graphics
    folder_name = create_unique_folder()

    # Загружаем только валидные листы
    valid_sheets = load_valid_sheets(file_path)
    if valid_sheets:
        print("Валидные листы:", ', '.join(valid_sheets) if valid_sheets[0] else "CSV файл")

        # Параметры генерации диаграмм из parameters.txt.
        # В программе введены дефолт значения, которые могут быть изменены
        params = read_parameters_from_file('parameters.txt')
        standard_deviation = params.get('standard_deviation', 4)
        show_original_values = params.get('show_original_values', True)

        for name in valid_sheets:
            make_diagrams(name, file_path, folder_name, standard_deviation, show_original_values)
    else:
        print("Нет валидных листов для обработки.")

    print("Program has done. Thank you for using.")
    print("parserDiagram_27uvs.")


# Запуск программы
if __name__ == '__main__':
    main()