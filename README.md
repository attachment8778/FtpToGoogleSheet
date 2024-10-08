**Документация по использованию**


**Ссылка на видео гайд по установке: ** (просто повторять и все)

1. [Установка питона:](https://disk.yandex.ru/i/Cn2hcPnPcCNNSQ) 

2. [Настройка конфигурации и прочего](https://disk.yandex.ru/i/9FNjWcbReCsKBQ)


**Описание программы:  
**Данная программа предназначена для загрузки файла ‘Amount.xlsx’ (Можно указать название файла свое) с сервера FTP, и последующей конвертации файла в Json формат и переноса в гугл таблицы, в будущем возможно добавления информации из json в качестве страницы заказов (дашбоард).

**Настройка конфигурации:**

**Для того что бы настроить конфигурацию файла, вам необходимо:**

1. **Установить библиотеки через консоль / программу которое вы используете:**
```py
Pip install xlrd==1.2.0 keyboard gspread oauth2cl
```

**2\. Изменение конфигурации в файле cfg.json**
```text
ftp_host: Адрес FTP-сервера.

**ftp_port**: Порт FTP-сервера (по умолчанию 21, если поставить значение None).

**ftp_user**: Имя пользователя для подключения к FTP.

**ftp_password**: Пароль для подключения к FTP.

**ftp_path**: Путь к файлу на FTP-сервере.

**ftp_file**: Имя файла на FTP-сервере.

**json_file**: Имя выходного JSON-файла.

**nameimen**: Список заголовков столбцов для преобразования Excel в JSON.

**times**: Список времени запуска задач в формате HH:MM.

**google_id**: Идентификатор Google таблицы. (Получаем из ссылки к самой таблице)  
```
1. **Получение файла googlekey.json:**

**Шаг 1:** Создание аккаунта в Google

Если у вас еще нет аккаунта в Google, создайте его.

**Шаг 2:** Создание проекта в [Google Cloud Platform](https://console.cloud.google.com/)
```text
- Перейдите на Google Cloud Platform.
- Убедитесь, что вы авторизованы в своем аккаунте Google.
- Создайте новый проект:
- Нажмите на "Создать проект".
- Укажите имя проекта и нажмите "CREATE".
```
**Шаг 3:** Создание сервисного аккаунта

1\. Перейдите в раздел "Credentials" по ссылке.

2\. Нажмите "CREATE SERVICE ACCOUNT".

3\. Укажите имя аккаунта и нажмите "DONE".

**Шаг 4:** Создание ключа для сервисного аккаунта

- Нажмите на три вертикальные точки возле созданного аккаунта и выберите "Manage keys".
- Нажмите "Add key" и выберите "Create new key".

3\. Выберите формат JSON и нажмите "CREATE".

4\. Скачайте JSON файл, и переименуйте его в googlekey.json.

5\. **Активируем Google Sheets API. Для этого переходим по ссылке** [Google Sheets API](https://console.cloud.google.com/marketplace/product/google/sheets.googleapis.com) и нажимаем на "ENABLE". Обратите внимание. Под каждый сервисный аккаунт необходимо выполнять активацию плагина **Google Sheets API**

**4\. Запуск программы:**

**Для запуска программы используйте следующие команды в консоле:  
```py
cd (путь к папке с ее указанием)  
py file.py
```

**Основные функции:**
```text
- **loadfromftp**: Загружает файл с FTP-сервера.
- **xltojs**: Преобразует данные из Excel в JSON.
- **update_google_sheet**: Обновляет данные в Google таблице.
- **service_run**: Основная функция, которая запускает сервис и следит за изменениями в конфигурационном файле.
```
**Примечания:**
```text
- Программа автоматически перезагружает конфигурацию при изменении cfg.json.
- Для немедленного запуска задач нажмите клавишу "R".
```
