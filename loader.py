import yadisk
import os
import shutil
import time
from tqdm import tqdm
from decouple import config

# функция для загрузки файлов
def loader():
    list_of_files = []
    load_path = ''

    # получение id, номера клиента и токена 
    client_id = config('client_id', default='')
    client_secret = config('client_secret', default='')
    ya_token = config('ya_token', default='')

    # получаем доступ на Я.Диске передаем id, номер клиента и токен. Хранятся в файле cfg.py
    y = yadisk.YaDisk(client_id, client_secret, ya_token)

    # проверяем токен
    while not y.check_token():
        print('Токен не прошел проверку, проверьте правильность токена')
    print('Проверка токена прошла успешно')
    
    # формируем лист со всем содержимым лежащим в директории 'lab' Я.Диска. 
    try:
        for el in y.listdir('lab'):
            if el['path'].endswith('.xlsx'):
                list_of_files.append(el['path'].split(':')[1])
    except Exception as e:
        print(f"Произошла ошибка: {e}")

    # Создаем директорию на локальной машине для хранения файлов  
    load_path = 'C:/loaded_files/'
    if not os.path.exists(load_path):
        os.mkdir(load_path)
    os.chdir(load_path)

    print('Локальная директория для хранения файлов: ', load_path)

    # Сообщение пользователю для понимания
    print("Выполняется загрузка файлов, пожалуйста подождите")

    # Скачивание файлов с Я.Диска
    for file in tqdm(list_of_files, dynamic_ncols=True):
        success = False
        attempts = 0
        max_attempts = 3  # максимальное количество попыток
        timeout = 5       # время ожидания в секундах между попытками
        # выполняем пока неуспешно и непревышено кол-во попыток
        while not success and attempts < max_attempts:
            try:
                y.download(file, os.path.join(load_path, file.split('/')[-1]))
                success = True
            except Exception as e:
                attempts += 1
                print(f"Ошибка при скачивании файла {file}: {e}")
                print(f"Попытка {attempts} из {max_attempts}. Повторная попытка через {timeout} секунд...")
                time.sleep(timeout)

        if not success:
            print(f"Файл {file} не удалось скачать после {max_attempts} попыток.")

    print("Загрузка файлов завершена")
    
    return list_of_files, load_path
    
# функция удаления файлов и директории
def delete_directory(path):
    # Проверяем, существует ли директория
    if os.path.exists(path):
        # Удаляем все файлы в директории
        for root, dirs, files in os.walk(path):
            for file in files:
                os.remove(os.path.join(root, file))

        # Переходим на уровень выше текущей директории
        os.chdir("..")  

        # Удаляем директорию
        shutil.rmtree(path)
        print(f"Директория '{path}' была успешно удалена вместе со всем содержимым.")
    else:
        print(f"Директория '{path}' не найдена.")
