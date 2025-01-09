import os
import shutil

import pandas as pd

from create_tables import generate_main_table, generate_no_matches_table
from excel_files import save_to_excel
from main_utils import clear_processed_folder
from settings import processed_folder
from main_utils import get_files_from_folder


def main_function(parent_folder):
    """Основаная функция программы, отвечающая за загрузку файлов, формирование таблиц и сохранение их в формате Excel"""

    # Обеспечивает работу программы пока в папке "не обработано" есть хоть одна папка с заявкой
    while True:
        # Если файл заявки и список счетов были получены, запускаем обработку файлов
        try:
            file, files_list = get_files_from_folder(parent_folder)  # Получаем файл заявки и список счетов
            if file and files_list:

                df1 = pd.read_excel(file)   # Загрузка файла заявки

                results_df = generate_main_table(df1, files_list)   # Формирование таблицы 1
                no_matches_df = generate_no_matches_table(files_list, results_df=results_df)   # Формирование таблицы 2

                # Формирование пути для сохранения итоговой таблицы в нужную папку с номером заявки
                folder_path = os.path.dirname(file)
                result_file_path = os.path.join(folder_path, "итоговая-таблица.xlsx")

                # Возможность сохранения итоговой таблицы несколько раз
                counter = 1
                while os.path.exists(result_file_path):
                    result_file_path = os.path.join(folder_path, f"итоговая-таблица ({counter}).xlsx")
                    counter += 1

                save_to_excel(results_df, no_matches_df, result_file_path)   # Сохранение итоговой таблицы в формате Excel

                shutil.move(folder_path, processed_folder)   # Перемещение папки с номером заявки из "не обработано" в "обработано"

                clear_processed_folder(processed_folder)   # Удаление папки с номером заявки, созданной более 7 дней назад

        except TypeError:
            # если файлы не были получены (в папке "не обработано" больше нет папок с заявками),
            # останавливает работу программы
            break
