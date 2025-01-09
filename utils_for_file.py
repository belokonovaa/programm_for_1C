import os

import docx
import pandas as pd
import pdfplumber
import win32com.client


def clean_dataframe(df):
    """Очистка DataFrame"""
    df = df.dropna(how='any')  # Удаляем строки, где все значения NaN
    df = df.dropna(axis=1, how='any')  # Удаляем столбцы, где все значения NaN
    df = df.fillna('')  # Заменяем пустые значения на пустые строки
    df.columns = df.columns.str.strip()  # Убираем пробелы в заголовках
    df = df.apply(lambda col: col.str.strip() if col.dtype == "object" else col)  # Убираем пробелы в ячейках
    return df


def extract_pdf_data(file2):
    """
        Извлекает таблицу из pdf документа и преобразует ее в DataFrame, если она существует.
        В случае ее отсутствия возвращает пустой DataFrame.
    """

    with pdfplumber.open(file2) as pdf_file:  # Читаем pdf файл

        for page in pdf_file.pages:  # Проходим по каждой странице в файле
            table = page.extract_table()  # Извлекаем таблицу
            if table:  # Если таблица существует, преобразуем ее в DataFrame
                df = pd.DataFrame(table[:], columns=table[0])
                df = clean_dataframe(df) # проводим предобработку таблицы

            return df
        return pd.DataFrame()


def convert_xls_to_xlsx(file2):
    """Конвертируем файл .xls в .xlsx"""
    new_file = pd.read_excel(file2, engine='xlrd')  # Чтение .xls файла
    new_file_path = file2.replace('.xls', '_converted.xlsx')  # Генерация нового пути для .xlsx файла
    new_file.to_excel(new_file_path, index=False, engine='openpyxl')  # Сохранение в .xlsx
    return pd.read_excel(new_file_path, engine='openpyxl')  # Чтение .xlsx файла после конвертации


def extract_docx_data(file2):
    """ Извлекает таблицу из Word документа и преобразует её в DataFrame."""

    doc = docx.Document(file2)  # Читаем Word файл

    for table in doc.tables:  # Ищем таблицы в документе
        table_data = []

        # Проверяем первую строку таблицы на наличие нужных колонок
        first_row = [cell.text.strip().lower() for cell in table.rows[0].cells]

        # Проверка на наличие столбцов "Товар" и "Цена"
        if any('товар' in column for column in first_row) and any('цена' in column for column in first_row):
            # Обрабатываем строки таблицы
            for row in table.rows:
                row_data = [cell.text.strip() for cell in row.cells]
                table_data.append(row_data)

            # Преобразуем таблицу в DataFrame
            df = pd.DataFrame(table_data[:], columns=table_data[0])  # Первая строка — это заголовки

            # Применяем предобработку DataFrame
            df = clean_dataframe(df)
            return df


def convert_doc_to_docx(file2):
    """Конвертируем Word-файл .doc в .docx"""

    # Преобразуем путь .doc файла в абсолютный (для возможности загрузки файла, имеющего в названии более 1 слова)
    doc_path = os.path.abspath(file2)
    base_name = os.path.splitext(doc_path)[0]
    docx_path = base_name + ".docx"  # Генерация нового пути для .docx файла

    # Открываем Word файл и конвертируем .doc в .docx
    word = win32com.client.Dispatch("Word.Application")
    doc = word.Documents.Open(doc_path)
    doc.SaveAs(docx_path, FileFormat=16)  # 16 — формат .docx
    doc.Close()
    word.Quit()

    df = extract_docx_data(docx_path)  # Обрабатываем файл как .docx и извлекаем таблицу

    return df


def load_data(file2):
    """Функция для загрузки любого типа файла"""

    file_path = os.path.splitext(file2)[1].lower()

    if file_path == '.pdf':  # если загружаемый файл имеет расширение .pdf
        return extract_pdf_data(file2)

    elif file_path == '.xlsx':  # если загружаемый файл имеет расширение .xlsx
        return pd.read_excel(file2)

    elif file_path == '.xls':   # если загружаемый файл имеет расширение .xls
        return convert_xls_to_xlsx(file2)

    elif file_path == '.docx':   # если загружаемый файл имеет расширение .docx
        return extract_docx_data(file2)

    elif file_path == '.doc':   # если загружаемый файл имеет расширение .doc
        return convert_doc_to_docx(file2)

