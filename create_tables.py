import os

import pandas as pd

from main_utils import get_columns, find_differences, find_best_match
from utils_for_value import find_price_difference, find_unit_difference, find_quantity_difference
from utils_for_file import load_data


def generate_main_table(df1, files_list):
    """Генерирует основную таблицу с результатами сопоставления товаров, цен и поставщиков"""

    results = []  # список для хранения всех полученных значений (будет использоваться для создания итоговой таблицы)

    # Проходим циклом по столбцам файла заявки для извлечения данных
    for index, row in df1.iterrows():
        item = row['Номенклатура']  # Строка из столбца 'Номенклатура'
        supplier_from_file1 = row['Поставщик']  # Строка из столбца 'Поставщик'
        price_from_file1 = row['Цена']   # Строка из столбца 'Цена'
        unit_from_file1 = row['Ед. изм.']   # Строка из столбца 'Единица измерения'
        quantity_from_file1 = row['Количество']   # Строка из столбца 'Количество'

        # Переменные для отслеживания информации
        best_match = ''   # лучшее совпадение номенклатуры
        best_score = 0   # процент совпадения
        best_supplier = None   # лучший поставщик
        best_price = None   # лучшая цена
        best_unit = None   # единица измерения лучшего совпадения
        best_quantity = None   # количество лучшего совпадения
        all_prices = []  # Список для всех цен
        suppliers = []  # Список для всех поставщиков
        items = []  # Список всех номенклатур
        units = []  # Список всех единиц измерения
        quantities = []  # Список всех значений "количество"
        differing_string = ""  # Инициализация переменной для столбца "несоответствие в тексте"

        # Проходим по каждому файлу счета
        for file2 in files_list:

            df2 = load_data(file2)  # Загружаем файл счета и извлекаем таблицу

            columns = get_columns(df2)  # Ищем нужные столбцы в извлеченной таблице

            # Смотрим какое название колонки цены прописано в таблице
            # Если "цена" - извлекаем данные из этого столбца, если "цена вкл. НДС" - берем данные из этого столбца.
            price_column = None
            if 'Цена вкл.НДС' in columns:
                price_column = columns['Цена вкл.НДС']
            elif 'Цена' in columns:
                price_column = columns['Цена']

            # Находим максимально похожее наименование из второго файла
            match, score, index = find_best_match(item, df2[columns['Товары']])

            # Если совпадение лучше, обновляем переменные
            if score > best_score and score >= 79:  # Порог для совпадений 79% для высокой точности
                best_match = match
                best_score = score
                best_supplier = os.path.basename(file2)
                price_values = df2[df2[columns['Товары']] == best_match][price_column].values

                # Обрабатываем случай, когда столбца "Единица измерения" нет в таблице
                if 'Единица' in columns:
                    best_unit_from_file2 = df2[df2[columns['Товары']] == best_match][columns['Единица']].values
                else:
                    best_unit_from_file2 = ["не указана"]

                best_quantity_from_file2 = df2[df2[columns['Товары']] == best_match][columns['Количество']].values

                if len(price_values) > 0:
                    price = price_values[0]
                    # Если цена для текущего совпадения меньше уже найденной, обновляем
                    if best_price is None or price < best_price:
                        best_price = price

                best_unit = best_unit_from_file2[0]
                best_quantity = best_quantity_from_file2[0]

            # Обработка случая, когда на одну позицию есть больше одного совпадения
            if score >= 79:
                matching_rows = df2[df2[columns['Товары']] == match]  # Все строки с этим товаром
                for _, row in matching_rows.iterrows():
                    price = row[price_column]
                    item_name = row[columns['Товары']]

                    # Если в таблице есть параметр "единица измерения" - берем ее, если нет - пишем "не указана"
                    if 'Единица' in columns:
                        unit_from_file2 = row[columns['Единица']]
                    else:
                        unit_from_file2 = ["не указана"]

                    quantity_from_file2 = row[columns['Количество']]
                    all_prices.append(price)
                    suppliers.append(os.path.basename(file2))
                    items.append(item_name)
                    units.append(unit_from_file2)
                    quantities.append(quantity_from_file2)

                    # Находим различие между строкой номенклатуры и товаром из счета
                    differing_string = find_differences(item, best_match)

        price_difference = find_price_difference(price_from_file1, best_price)  # Формирование столбца "Несоответствие в цене"
        unit_difference = find_unit_difference(unit_from_file1, best_unit)   # Формирование столбца "Несоответствие в единице измерения"
        quantity_difference = find_quantity_difference(quantity_from_file1, best_quantity)  # Формирование столбца "Несоответствие в количестве"

        # Формирование столбца "Дубляжи и схожие позиции"
        all_prices_str = "\n\n ".join(
            [f"{p} - {item}. \nКоличество - {quantity}. \nЕд. изм. - {unit} \n({s})" for item, p, s, unit, quantity in
             zip(items, all_prices, suppliers, units, quantities)][1:])

        # Добавляем строку в итоговую таблицу
        results.append([item, price_from_file1, quantity_from_file1, unit_from_file1, supplier_from_file1,
                        best_match, best_price, best_quantity, best_unit, best_supplier, all_prices_str,
                        differing_string, price_difference, quantity_difference, unit_difference])

    # Формируем итоговую таблицу
    result_df = pd.DataFrame(results,
                             columns=[' Наименование номенклатуры ', ' Цена ', ' Количество ',' Ед.изм. ',' Поставщик ',
                                      'Наименование номенклатуры', 'Цена', 'Количество', 'Ед.изм.', 'Поставщик',
                                      'Дубли и схожие позиции', 'в тексте', 'в цене', 'в количестве', 'в ед.изм.'])

    # Добавляем столбец '№' с индексами, начиная с 1
    result_df['№'] = result_df.index + 1

    # Перемещаем столбец '№' на первую позицию
    result_df = result_df[[
        '№', ' Наименование номенклатуры ', ' Цена ', ' Количество ', ' Ед.изм. ', ' Поставщик ',
        'Наименование номенклатуры', 'Цена', 'Количество', 'Ед.изм.', 'Поставщик',
        'Дубли и схожие позиции', 'в тексте', 'в цене', 'в количестве', 'в ед.изм.']]

    # Добавляем пустые столбцы для визуального деления таблицы
    result_df.insert(6, '1', '')  # Пустой столбец после 'Поставщик (заявка)'
    result_df.insert(12, '2', '') # Пустой столбец после 'Поставщик (счет)'

    return result_df


def process_no_matches(all_tovary, best_matches):
    """Обрабатывает товары, которые не были найдены среди лучших совпадений"""

    no_matches = []  # Список позиций без совпадения

    # Проходим циклом по всем позициям из файлов счетов
    for item in all_tovary:
        name, price, quantity, unit, file_name = item[0], item[1], item[2], item[3], item[4]  # Распаковываем данные

        # Проверяем на наличие в best_matches. Если эта позиция не числится в списке лучших совпадений,
        # добавляем ее в список позиций без совпадения no_matches
        if name not in best_matches:
            no_matches.append([name, price, quantity, unit, file_name])

    return no_matches


def generate_no_matches_table(files_list, results_df):
    """Генерирует таблицу с позициями без совпадения (процент совпадения которых меньше 80%) """

    all_tovary = []  # Список всех позиций из файлов счетов
    best_matches = results_df['Наименование номенклатуры'].tolist()  # Список лучших совпадений

    # Проходим циклом по списку счетов
    for file2 in files_list:
        file_name = os.path.basename(file2)

        df2 = load_data(file2)  # Загружаем файл счета и извлекаем таблицу
        columns = get_columns(df2)  # Ищем нужные столбцы в извлеченной таблице

        # Смотрим какое название колонки цены прописано в таблице
        # Если "цена" - извлекаем данные из этого столбца, если "цена вкл. НДС" - берем данные из этого столбца.
        price_column = None
        if 'Цена вкл.НДС' in columns:
            price_column = columns['Цена вкл.НДС']
        elif 'Цена' in columns:
            price_column = columns['Цена']

        quantity_column = columns.get('Количество', None)  # Получаем данные для столбца "Количество"
        unit_column = columns.get('Единица', None)  # Получаем данные для столбца "Ед.изм."

        # Если данные существуют, извлекаем их из соответствующих столбцов
        if price_column and quantity_column and unit_column:
            df2_cleaned = df2[[columns['Товары'], price_column, quantity_column, unit_column]].dropna()

            # Исключаем строки, которые содержат текст, не относящийся к нужной нам таблице
            df2_cleaned = df2_cleaned[
                ~df2_cleaned[columns['Товары']].str.contains(
                    'Внимание|Оплата|Уведомление|работы|услуги|Сч. №|БИК|Товар|Итого по счету:|Доставка',
                    na=False)]

            df2_cleaned['Название файла'] = file_name  # Получаем данные для столбца "Поставщик"

            all_tovary.extend(df2_cleaned.values.tolist())  # Добавляем отфильтрованные товары в список

    no_matches = process_no_matches(all_tovary, best_matches)  # Получаем список позиций без совпадений

    # Формируем столбцы в таблице
    columns = pd.MultiIndex.from_tuples([("Позиции без соответствия", "Наименование номеклатуры"),
                                         ("Позиции без соответствия", "Цена"),
                                         ("Позиции без соответствия", "Количество"),
                                         ("Позиции без соответствия", "Ед. изм."),
                                         ("Позиции без соответствия", "Название файла")])

    no_matches_df = pd.DataFrame(no_matches, columns=columns)  # Генерируем таблицу с позициями без совпадений

    no_matches_df.index = range(1, len(no_matches_df) + 1)  # Добавляем нумерацию в таблицу

    return no_matches_df
