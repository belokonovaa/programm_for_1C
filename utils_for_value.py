

def convert_to_float(value):
    """Функция для предобработки чисел (столбцы "Цена" и "Количество")"""

    # Если входные данные представлены в виде строки
    if isinstance(value, str):
        value_str = value.replace(" ", "").replace(",", ".")  # Удаляем пробелы и заменяем запятую на точку
        try:
            return float(value_str)  # Возвращаем тип float, в случае невозможности - значение None
        except ValueError:
            return None

    # Если входные данные представлены в виде числа
    if isinstance(value, (int, float)):
        return float(value)  # Возвращаем тип float, в случае невозможности - значение None
    return None


def find_price_difference(price_from_file1, best_price):
    """
        Функция для заполнения столбца 'Несоответствие в цене', сравнивает цену из заявки и цену из счета
        и выводит разницу, если она есть, или оставляет ячейку пустой.
    """

    # Проводим предобработку полученных цен
    price_from_file1 = convert_to_float(price_from_file1)   # Цена из заявки
    best_price = convert_to_float(best_price)   # Цена из счета

    # Если цены существуют, проводим их сравнение
    if price_from_file1 is not None and best_price is not None:
        if isinstance(price_from_file1, (int, float)) and isinstance(best_price, (int, float)):
            if price_from_file1 != best_price:  # Если цены не равны, указываем разницу
                return f' {round(price_from_file1 -  best_price, 2)} \n\n{price_from_file1} / {best_price}'
        else:
            return None   # Если цены равны, оставляем ячейку пустой


def normalize_unit(unit):
    """ Функция для предобработки строк ( столбец 'Единица измерения') """

    # Словарь с данными единиц измерения
    # (формирует связь, чтобы программа "м" и "метр" или "шт" и "штука" считывала за одну единицу измерения)
    unit_synonyms = {
        "м": "метр", "метры": "метр", "метр": "метр", "Метр": "метр", "мт": "метр",
        "шт": "штука", "штука": "штука", "Штука": "штука", "штуки": "штука",
        "уп": "упаковка", "упак": "упаковка", "упаковка": "уп"
    }

    # Если полученное значение - строка, приводим все слова в ней к нижнему регистру
    if isinstance(unit, str):
        return unit_synonyms.get(unit.lower(), unit.lower())
    return str(unit)  # Если это не строка, преобразуем в строку


def find_unit_difference(unit_from_file1, best_unit):
    """
        Функция для заполнения столбца 'Несоответствие в ед.изм.', сравнивает ед.изм. из заявки и ед.изм. из счета
        и выводит разницу, если она есть, или оставляет ячейку пустой.
    """

    if unit_from_file1 is not None and best_unit is not None:
        # Проводим предобработку полученных единиц измерения
        normalized_unit_from_file1 = normalize_unit(unit_from_file1)   # Ед.изм. из заявки
        normalized_best_unit = normalize_unit(best_unit)   # Ед.изм. из счета

        # Сравниваем значения, если они равны - оставляем ячейку пустой
        if normalized_unit_from_file1 == normalized_best_unit:
            return None
        else:
            return f' {unit_from_file1} / {best_unit}'   # Если значения не равны, выводим разницу


def find_quantity_difference(quantity_from_file1, best_quantity):
    """
        Функция для заполнения столбца 'Несоответствие в количестве', сравнивает количество из заявки и количество из счета
        и выводит разницу, если она есть, или оставляет ячейку пустой.
    """

    # Проводим предобработку полученных значений количества
    quantity_from_file1 = convert_to_float(quantity_from_file1)   # Количество из заявки
    best_quantity = convert_to_float(best_quantity)   # Количество из счет

    # Если значения существуют, проводим их сравнение
    if quantity_from_file1 is not None and best_quantity is not None:
        if isinstance(quantity_from_file1, (int, float)) and isinstance(best_quantity, (int, float)):
            if quantity_from_file1 != best_quantity:   # Если значения не равны, выводим разницу
                return f' {best_quantity - quantity_from_file1} \n\n {quantity_from_file1} / {best_quantity}'
        else:
            return None   # Если значения равны, оставляем ячейку пустой