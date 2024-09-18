# libs   
import os
import json
import time
import datetime
import ftplib  
import os
import json
import xlrd
import time
import keyboard
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta

# Загрузка данных из cfg.json
config_path = os.path.join(os.path.dirname(__file__), 'cfg.json')
with open(config_path, 'r', encoding='utf-8') as config_file:
    data = json.load(config_file)

ftp_host = data['ftp_host']
ftp_port = data['ftp_port']
ftp_user = data['ftp_user']
ftp_password = data['ftp_password']
ftp_path = data['ftp_path']
ftp_file = data['ftp_file']
json_file = data['json_file']
nameimen = data['nameimen']
times = data['times']
google_id = data['google_id']

# func connect and load file ftp  
err = 0
def loadfromftp(host, port, usname, passwd, path, file):   
    try:    
        ftp = ftplib.FTP()
        if port is None or not port:
            port = 21
        print(f'Данные подключения: \n host: {host}\n user: {usname}\n Path: {path} \n loaded file: {file}\n current port: {port}')
        ftp.connect(host, port)
        ftp.set_pasv(True)
        ftp.login(user=usname, passwd=passwd)
        #go to path
        ftp.cwd(path)
        
        # Download file
        with open(os.path.join(os.path.dirname(__file__), file), 'wb') as local_file:
            ftp.retrbinary(f'RETR {file}', local_file.write)
        print(f"Файл {file} успешно загружен.")
        ftp.quit()
    except ftplib.error_perm as perm_error:
        raise SystemExit(f"Ошибка разрешений: {perm_error}")
    except Exception as error:
        raise SystemExit(f"Общая ошибка: {error}")

def xltojs(excel_file, json_file, nameimen):
    excel_file_path = os.path.join(os.path.dirname(__file__), excel_file)
    workbook = xlrd.open_workbook(excel_file_path)
    sheet = workbook.sheet_by_index(0)
    data = []
    for row_idx in range(1, sheet.nrows):
        row = {}
        for col_idx, header in enumerate(nameimen):
            value = sheet.cell_value(row_idx, col_idx)
            if value == "":
                value = "Отсутствует"
            elif isinstance(value, float) and value.is_integer():
                value = int(value) 
            row[header] = value
        data.append(row)
    with open(os.path.join(os.path.dirname(__file__), json_file), 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
        print(f"Файл {json_file} успешно создан.")




def service_run():
    def retry_function(func, max_retries, *args):
        for attempt in range(max_retries):
            try:
                result = func(*args)
                return result
            except Exception as error:
                print(f'Ошибка: {error}')
            if attempt < max_retries - 1:
                print(f'Повторная попытка {attempt + 1}...')
            time.sleep(1)
        return False

    def load_config():
        config_path = os.path.join(os.path.dirname(__file__), 'cfg.json')
        with open(config_path, 'r', encoding='utf-8') as config_file:
            return json.load(config_file)

    def run_service():
        last_modified = os.path.getmtime('cfg.json')
        config = load_config()

        def check_config_changes():
            nonlocal last_modified, config
            current_modified = os.path.getmtime('cfg.json')
            if current_modified > last_modified:
                print("Обнаружены изменения в cfg.json. Перезагрузка конфигурации...")
                config = load_config()
                last_modified = current_modified
                return True
            return False

        def run_tasks():
            nonlocal config
            if err == 1:
                print('Произошла ошибка. Дальнейшие действия остановлены.')
                return
            ftp_result = retry_function(loadfromftp, 3, config['ftp_host'], config['ftp_port'], 
                                        config['ftp_user'], config['ftp_password'], 
                                        config['ftp_path'], config['ftp_file'])
            if ftp_result is not False:
                conversion_result = retry_function(xltojs, 3, config['ftp_file'], 
                                                   config['json_file'], config['nameimen'])
                if conversion_result is False:
                    print('Проблема при преобразовании.')
                else:
                    update_result = retry_function(update_google_sheet, 3)
                    if update_result is False:
                        print('Проблема при обновлении Google таблицы.')
                    else:
                        print('Все задачи выполнены успешно.')
            else:
                print('Ошибка при загрузке файла с FTP.')

        while True:
            if check_config_changes():
                print("Конфигурация обновлена.")

            now = datetime.now()
            next_run_time = None
            times_sorted = sorted(config['times'])
            for t in times_sorted:
                run_time = datetime.strptime(t, '%H:%M').replace(year=now.year, month=now.month, day=now.day)
                if run_time > now:
                    next_run_time = run_time
                    break
            if not next_run_time:
                next_run_time = datetime.strptime(times_sorted[0], '%H:%M').replace(year=now.year, month=now.month, day=now.day) + timedelta(days=1)
            
            print(f'Сервис запущен. Следующая работа будет в {next_run_time.strftime("%H:%M")}')
            print('Нажмите "R" для немедленного запуска задач')

            time_to_wait = (next_run_time - now).total_seconds()
            start_time = time.time()

            while time.time() - start_time < time_to_wait:
                if check_config_changes():
                    print("Конфигурация обновлена. Пересчет времени следующего запуска...")
                    break
                if keyboard.is_pressed('r'):
                    print('Запуск задач по требованию...')
                    run_tasks()
                    break
                time.sleep(0.1)

            if time.time() - start_time >= time_to_wait:
                run_tasks()

    run_service()

def update_google_sheet():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('googlekey.json', scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(google_id).sheet1
    workbook = xlrd.open_workbook('amount.xlsx')
    worksheet = workbook.sheet_by_index(0)
    data = []
    for row in range(worksheet.nrows):
        data.append(worksheet.row_values(row))
    sheet.clear()
    sheet.update(data)
    print("Данные успешно обновлены в Google таблице.")

    
service_run()
#loadfromftp(ftp_host, ftp_port, ftp_user, ftp_password, ftp_path, ftp_file)
#xltojs(ftp_file, json_file, nameimen)







