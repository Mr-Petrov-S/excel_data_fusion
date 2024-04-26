import pandas as pd
from loader import loader, delete_directory  
from preprocessors import two_lists, preprocessor_sop, preprocessor_zno, save_files

def main():

    # Запускаем функцию загрузки файлов
    list_of_files, load_path = loader()

    # Формируем два списка по направлениям ЗНО и СОП
    list_of_sop, list_of_zno = two_lists(list_of_files)

    # Запускаем обработчик для файлов СОП
    table_sop_1, table_sop_2, table_sop_3 = preprocessor_sop(list_of_sop)

    # Запускаем обработчик для файлов ЗНО
    table_zno = preprocessor_zno(list_of_zno)

    # Запускаем функцию для сохранения файлов
    save_files(table_sop_1, "table_sop_1")
    save_files(table_sop_2, "table_sop_2")
    save_files(table_sop_3, "table_sop_3")
    save_files(table_zno, 'table_zno')
    
    # предлагаем пользователю удалить скачанные файлы с локальной машины
    print("Вы хотите удалить скачанные файлы и директорию? (y/n): ")

    user_input = input()

    while user_input.lower() not in ['y', 'n']:
        user_input = input("Пожалуйста, нажмите 'y' или 'n'. ")
    
    # если пользователь ответил да, вызываем функцию удаления директории и файлов
    if user_input == 'y':
        delete_directory(load_path)
    elif user_input == 'n':
        print('Обработка закончена, хорошего дня!')


if __name__ == "__main__":
    main()