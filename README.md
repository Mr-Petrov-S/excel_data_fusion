# Добро пожаловать!

Этот скрипт написан для обработки определенных файлов в рамках поставленной задачи и не является универсальным.
Он был написан для проекта обработки данных по онкопомощи, в рамках социально значимой инициативы создания единой базы знаний о динамике и состоянии онкопомощи в РФ.
Он позволяет объединить произвольное количество excel файлов определенной наполненности в 4 общих, разделенных по смыслу файла. 


## Описание работы: 
- После запуска, будет сформирован список файлов хранящийся в директории 'lab на Яндекс Диске
- Далее будет выполнено скачивание этих файлов на локальную машину в директорию 'C:/loaded_files/', если такой директории нет, она будет создана. 
Так как скачивание может быть не бытрым процессом, срипт оснащен прогресс баром для отслеживания, в случае ошибки будет выдано соответствующее описание. 
- После скачивания будут сформированы списки для обработки по направлениям "Состояние онко помощи"(далее СОП) и "Злокачественные новообразования"(далее ЗНО)
- Далее будет запущена обработка списка таблиц СОП, в ходе выполнения которой в терминал будут выводится названия файлов которые находятся в обработке, а также номера листов.
По окончании включится следующий обработчик для файлов ЗНО. Он также оснащен аналогичными выводами в терминал. 
- После завершения обработки запустится сохранение этих файлов на локальную машину в директорию "C:/processed_files/". В эту папку будут добавлены итоговые объединенные таблицы. 
- После сохранения вы можете выбрать удалить скачанные ранее с Яндекс Диска файлы с локального устройства или оставить их там. 
На этом работа программы прекращена. 

## Краткое описание файлов:
main.py - главный исполнительный файл
preprocessors.py - файл содержащий функции обработчиков, сохранения файлов и формирования списков СОП/ЗНО из перечня скачанных файлов
loader.py - файл в котором храняытся функции скачивания с Яндекс Диска на локальное устройство, а также функция удаления с локального устройства
cancer_loc.py - файл со списками болезней присущих только определенному полу
regions.py - файл со словарем федеральных округов.
env.example - пример .env файла необходимого для работы.


**Для корректной работы требуется .env файл, который будет содержать валидные id и токен Яндекс Диска. Пример подобного файла - env.example.**

### Ниже описаны шаги для получения токена для работы с Яндекс.Диском

Для получения токена необходимо:
1. Зайти на сайт для получения доступа к ресурсам Яндекса: https://oauth.yandex.ru
2. Авторизоваться на сайте под своим аккаунтом Яндекс ID
3. На главной странице нажать «Создать приложение»
4. На странице создания приложения необходимо удалить «id» из пути, оставив только https://oauth.yandex.ru/client/new/ в адресной строке браузера
5. Нас переместит на страницу создания приложения, в которой можно выбрать необходимые права для доступа к Яндекс.Диску
    - В разделе «Общие данные» в поле ввода «Название вашего сервиса» вводим название для создаваемого приложения (любое);
    - В разделе «Платформы приложения» выбираем «Веб-сервисы»;
    - Щелкаем мышью на появившееся поле ввода «Redirect URI» и во всплывшей подсказке щелкаем мышью на «Подставить URL для отладки».
6. Во вкладке «Доступ к данным» в поле ввода вбиваем слово «disk» и в выпадающем меню выбираем доступ «Чтение всего диска»
7. В поле «Почта для связи» указываете свою почту
8. На странице приложения нам понадобится «ClientID» – копируем его. Открываем новое окно браузера и вбиваем в него следующий адрес:
    https://oauth.yandex.ru/authorize?response_type=token&client_id=
    После знака «равно» добавляем ClientID вашего приложения, итоговая ссылка будет выглядеть примерно так:
    https://oauth.yandex.ru/authorize?response_type=token&client_id=с0000ff0000000c0ba00000e00e00000
    Нажимаем «Enter» на клавиатуре и, если все сделано правильно, то нас переместит на страницу с подтверждением получения доступов для стороннего приложения.
9. Входим в свой аккаунт После чего нас перемещает на страницу с токеном, который необходимо сохранить в файл или запомнить – он понадобится для работы с диском.
