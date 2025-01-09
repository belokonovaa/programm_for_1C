import time

from main import main_function
from settings import parent_folder, pause_time


if __name__ == "__main__":
    while True:
        main_function(parent_folder)  # основная функция программы
        time.sleep(pause_time)  # автоматический запуск программы