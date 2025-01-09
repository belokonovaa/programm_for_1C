import os
import re
import shutil
import time

from fuzzywuzzy import process, fuzz


def find_differences(str1, str2):
    """
        Возвращает различия между двумя строками и возвращает слова, которые есть во второй строке, но отсутствуют в первой.
    """

    # Проводим предобработку строк - Убираем все символы, кроме букв, цифр и пробела
    str1_cleaned = re.sub(r'[^\w\s]', '', str1)
    str2_cleaned = re.sub(r'[^\w\s]', '', str2)

    # Разбиваем строки на слова
    words1 = set(str1_cleaned.split())  # первую строку преобразуем в множество
    words2 = str2_cleaned.split()

    # Ищем слова, которые есть во второй строке, но отсутствуют в первой
    differences = [word for word in words2 if word not in words1]

    return ' '.join(differences)


def find_header_row(df):
    """
    Ищет строку с заголовками в DataFrame, где присутствуют ключевые слова "товар" и "цена" (или "цена вкл.ндс").
    Возвращает индекс строки с заголовками, или None, если такая строка не найдена.
    """
    for i, row in df.iterrows():
        if any('товар' in str(val).lower() for val in row) and any(
                ('цена' in str(val).lower() or 'цена вкл.ндс' in str(val).lower()) for val in row):
            return i
    return None


def find_columns(df, header_row):
    """
    Определяет столбцы в найденной строке, которые содержат заданные ключевые слова.
    Возвращает словарь, где ключи - это имена столбцов, а значения - их названия.
    """
    df.columns = df.iloc[header_row]  # Устанавливаем заголовки
    df = df.drop(header_row)  # Убираем строку с заголовками из данных, чтобы она не использовалась при поиске совпадений в строках

    # Создаем словарь
    possible_columns = {}

    for col in df.columns:
        if 'товар' in str(col).lower():  # Ищем столбец с товарами
            possible_columns['Товары'] = col
        if 'цена' in str(col).lower():  # Ищем столбец с ценой
            possible_columns['Цена'] = col
        if 'вкл' in str(col).lower() and 'ндс' in str(col).lower():  # Ищем столбец с ценой вкл. НДС
            possible_columns['Цена вкл.НДС'] = col
        if 'ед' in str(col).lower():  # Ищем столбец с единицей измерения
            possible_columns['Единица'] = col
        if 'кол' in str(col).lower(): # Ищем столбец с количеством
            possible_columns['Количество'] = col

    return possible_columns


def get_columns(df):
    """Объединяет результат функций find_header_row и find_columns"""
    header_row = find_header_row(df)
    return find_columns(df, header_row)


def find_best_match(query, choices):
    """Находит лучшее соответствие для строки query (строка из заявки) в списке choices (строки из счетов)"""
    match, score, index = process.extractOne(query, choices, scorer=fuzz.token_set_ratio)
    return match, score, index


def get_files_from_folder(parent_folder):
    """Возвращает файл заявки и список файлов счетов"""
    application_file = None  # Переменная для файла заявки

    # Проходим по всем папкам в "необработано"
    for folder_name in os.listdir(parent_folder):
        folder_path = os.path.join(parent_folder, folder_name)

        # Проверяем, является ли это папкой
        if os.path.isdir(folder_path):
            current_invoice_files = []  # Список счетов для текущей папки

            # Проходим по всем файлам в папке
            for filename in os.listdir(folder_path):
                full_path = os.path.join(folder_path, filename)

                # Проверяем, является ли это файлом (и не папкой)
                if os.path.isfile(full_path):
                    # Если файл заявки, сохраняем его в переменную application_file
                    if filename.lower() == 'заявка.xlsx':  # Пример имени файла заявки
                        application_file = full_path
                    else:
                        # Все остальные файлы добавляем в список счетов текущей папки
                        current_invoice_files.append(full_path)

            return application_file, current_invoice_files
        return None, None


def clear_processed_folder(folder_path):
    """Очищает содержимое папки 'Обработано', если возраст вложенных папок более 7 суток (604 800 секунд)"""

    current_time = time.time()  # Получаем текущее время

    # Перебираем все элементы в папке
    for folder_name in os.listdir(folder_path):
        folder_path_to_check = os.path.join(folder_path, folder_name)

        # Вычисляем возраст папки по времени последнего изменения
        folder_age = current_time - os.path.getmtime(folder_path_to_check)

        # Если папка старше 7 суток (604800 секунд), удаляем её вместе с содержимым
        if folder_age > 604800:
            shutil.rmtree(folder_path_to_check)
