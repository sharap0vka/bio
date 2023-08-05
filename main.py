import dearpygui.dearpygui as dpg
import datetime
import re
import os
import os.path
import json
import configparser
from time import sleep
# from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
from pprint import pprint
import pyperclip
import dearpygui.demo as demo

config = configparser.ConfigParser()
config.read('config.ini')

DEFAULT_DATE = config['APP']['DEFAULT_DATE']
MID_DAY_HOUR = int(config['APP']['MID_DAY_HOUR'])
DRIVER_PATH = config['APP']['DRIVER_PATH']
LOG_PATH = config['APP']['LOG_PATH']
LOGIN_ACC = config['APP']['LOGIN_ACC']
PASSWORD_ACC = config['APP']['PASSWORD_ACC']

DATE_BIO8_START = datetime.datetime.strptime('19.07.2021', '%d.%m.%Y').date()

log = {}
BUFFER = []
RED_COLOR = (255, 0, 0, 255)
GREEN_COLOR = (255, 0, 0, 255)
YELLOW_COLOR = (255, 255, 0, 255)
ELEMENTS = []

dpg.create_context()

with dpg.font_registry():
    with dpg.font("./fonts/JetBrainsMonoNL-Regular.ttf", 20, default_font=True, id="Default font"):
        dpg.add_font_range_hint(dpg.mvFontRangeHint_Cyrillic)

class TableData():
    def __init__(self, events):
        self.events = events
    
    def get_count_all_events(self):
        return len(self.events)
    
    def get_count_all_abort(self):
        counter = 0
        for event in self.events:
            if event['id'] == None: counter += 1
        return counter
    
    def get_count_morn_abort(self):
        counter = 0
        for event in self.events:
            if event['id'] == None and get_timestamp(event['time']) == 'Утро': counter += 1
        return counter
    
    def get_count_even_abort(self):
        counter = 0
        for event in self.events:
            if event['id'] == None and get_timestamp(event['time']) == 'Вечер': counter += 1
        return counter

class Query():
    def __init__(self, shop, events):
        self.shop = shop
        self.events = events

    def get_start_date(self):
        min_date = self.events[0]['date']
        for event in self.events:
            if event['date'] < min_date:
                min_date = event['date']
        return min_date
    
    def get_end_date(self):
        max_date = self.get_start_date()
        for event in self.events:
            if event['is_day'] == 'Day':
                if event['date'] > max_date: max_date = event['date']
            if event['is_day'] == 'Night':
                if event['date'] + datetime.timedelta(days=1) > max_date:
                    max_date = event['date'] + datetime.timedelta(days=1)
        return max_date
    
class Response():
    def __init__(self, events, log):
        self.events = list.copy(events)
        self.log = dict.copy(log)

    def make(self):
        for event in self.events:
            worker = event['worker']
            event['confirmed_mark'] = []
            date = event['date']
            date2 = date + datetime.timedelta(days=1)
            shift = event['is_day']
            if shift == 'Day':
                if date.strftime("%d.%m.%Y") in self.log: # Есть ли дата в логе
                    if worker in self.log[date.strftime("%d.%m.%Y")]: # Ессть ли отметки сотрудника за дату
                        event['confirm'] = self.log[date.strftime("%d.%m.%Y")][worker]
                    else:
                        event['confirm'] = []
                        
                    if None in self.log[date.strftime("%d.%m.%Y")]:
                        event['abort'] = self.log[date.strftime("%d.%m.%Y")][None]
                    else:
                        event['abort'] = []
                else:
                    event['confirm'] = []
                    
                    event['abort'] = []
            if shift == 'Night':
                if date.strftime("%d.%m.%Y") in self.log: # Проверяем ПРИХОД
                    if worker in self.log[date.strftime("%d.%m.%Y")]:
                        event['confirm_day1'] = self.log[date.strftime("%d.%m.%Y")][worker]
                    else:
                        event['confirm_day1'] = []
                        
                    if None in self.log[date.strftime("%d.%m.%Y")]:
                        event['abort_day1'] = self.log[date.strftime("%d.%m.%Y")][None]
                    else:
                        event['abort_day1'] = []
                else:
                    
                    event['confirm_day1'] = []
                    event['abort_day1'] = []
                if date2.strftime("%d.%m.%Y") in self.log: # Проверяем УХОД
                    if worker in self.log[date2.strftime("%d.%m.%Y")]:
                        event['confirm_day2'] = self.log[date2.strftime("%d.%m.%Y")][worker]
                    else:
                        event['confirm_day2'] = []
                        
                    if None in self.log[date2.strftime("%d.%m.%Y")]:
                        event['abort_day2'] = self.log[date2.strftime("%d.%m.%Y")][None]
                        
                    else:
                        event['abort_day2'] = []
                else:
                    event['confirm_day2'] = []
                    
                    event['abort_day2'] = []

    def set_status(self, event):
        if 'status' in event:
            if event['status'] == 'not_found':
                return
        is_day = event['is_day']
        if 'bio8_fail' in event:
            event['status'] = 'fail_bio8'
            return
        if 'start_work_date' in event:
            if event['date'] < event['start_work_date']:
                event['status'] = 'fail_start_work_date'
                return
        if event['pass'] < 2:
            event['status'] = 'fail_pass_mark'
            return
        if is_day == 'Day':
            check_in = False
            check_out = False
            for mark in event['confirm']:
                if get_timestamp(mark) == 'Утро':
                    check_in = True
                    event['confirmed_mark'].append({
                        'date': event['date'].strftime("%d.%m.%Y"),
                        'time': mark,
                        'about': 'Приход распознан'
                    })
                if get_timestamp(mark) == 'Вечер':
                    check_out = True
                    event['confirmed_mark'].append({
                        'date': event['date'].strftime("%d.%m.%Y"),
                        'time': mark,
                        'about': 'Уход распознан'
                    })
            if check_in and check_out:
                event['status'] = 'succ_axapta'
                return
            if check_in == False or check_out == False:
                for mark in event['abort']:
                    if get_timestamp(mark) == 'Утро':
                        check_in = True
                    if get_timestamp(mark) == 'Вечер':
                        check_out = True
                if check_in and check_out:
                    event['status'] = 'succ_mark'
                    return
                if check_in and check_out == False:
                    event['status'] = 'fail_out'
                if check_in == False and check_out:
                    event['status'] = 'fail_in'
                if check_in == False and check_out == False:
                    event['status'] = 'fail_all'
        if is_day == 'Night':
            check_in = False
            check_out = False
            for mark in event['confirm_day1']:
                if get_timestamp(mark) == 'Вечер':
                    check_in = True
                    event['confirmed_mark'].append({
                        'date': event['date'].strftime("%d.%m.%Y"),
                        'time': mark,
                        'about': 'Приход распознан'
                    })
            for mark in event['confirm_day2']:
                if get_timestamp(mark) == 'Утро':
                    check_out = True
                    date2 = event['date'] + datetime.timedelta(days=1)
                    event['confirmed_mark'].append({
                        'date': date2.strftime("%d.%m.%Y"),
                        'time': mark,
                        'about': 'Уход распознан'
                    })
            if check_in and check_out:
                event['status'] = 'succ_axapta'
                return
            if check_in == False:
                for mark in event['abort_day1']:
                    if get_timestamp(mark) == 'Вечер':
                        check_in = True
            if check_out == False:
                for mark in event['abort_day2']:
                    if get_timestamp(mark) == 'Утро':
                        check_out = True
            if check_in and check_out:
                event['status'] = 'succ_mark'
                return
            if check_in and check_out == False:
                event['status'] = 'fail_out'
                return
            if check_in == False and check_out:
                event['status'] = 'fail_in'
            if check_in == False and check_out == False:
                event['status'] = 'fail_all'
        
    def calculate(self):
        for event in self.events:
            self.set_status(event)

def start_chrome(driver, shop_number, events, date_in, date_out):
    service = Service(driver)
    # service = Service(ChromeDriverManager().install())
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--headless=new')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.get(f'http://{LOGIN_ACC}:{PASSWORD_ACC}@srv-setmon-dcb/BioTools/Check.aspx')

    # Считаем кол-во отпечатков в системе
    radio = driver.find_element(By.ID, 'RadioButtonList2_0')
    radio.click()
    for event in events:
        search_mark_input = driver.find_element(By.ID, 'CR_Name')
        search_btn = driver.find_element(By.ID, 'CR_Search')
        search_mark_input.clear()
        search_mark_input.send_keys(event['worker'])
        search_btn.click()
        page_source = driver.page_source
        soup = BeautifulSoup(page_source, 'html.parser')
        try:
            rows = soup.find('table', {'id': 'GridView1'}).find_all('tr')
            cells = soup.find('table', {'id': 'GridView1'}).find_all('td')
            counter = 0
            counter_cell = 1
            if rows:
                for row in rows:
                    if str(row).find(event['worker']) >=0:
                        counter += 1
            if cells:
                for cell in cells:
                    if counter_cell % 5 == 0:
                        date_employment = datetime.datetime.strptime(cell.text, '%d.%m.%Y %H:%M:%S').date()
                        if date_employment < DATE_BIO8_START:
                            event['bio8_fail'] = True
                    counter_cell += 1
            sleep(1)
            event['pass'] = counter
        except Exception:
            event['status'] = 'not_found'

    # Сверяем дату устройства
    
    # for event in events:
    #     worker_axapta_input = driver.find_element(By.ID, 'St_TB')
    #     worker_axapta_button = driver.find_element(By.ID, 'St_Button')
    #     worker_axapta_input.clear()
    #     worker_axapta_input.send_keys(event['worker'])
    #     worker_axapta_button.click()
    #     page_source = driver.page_source
    #     soup = BeautifulSoup(page_source, 'html.parser')
    #     table = soup.find('table', {'id': 'GridView1'})
    #     td = None
    #     if table: td = table.find_all('td')
    #     if td:
    #         event['first_name'] = td[1].text
    #         event['last_name'] = td[3].text
    #         event['start_work_date'] = datetime.datetime.strptime(td[11].text, '%d.%m.%Y %H:%M:%S').date()
    #     sleep(1)
    
    # ВЫГРУЗКА
    log_shop_input = driver.find_element(By.ID, 'LogMag')
    log_shop_input.send_keys(shop_number)
    log_start_date = driver.find_element(By.ID, 'LogDate1')
    log_end_date = driver.find_element(By.ID, 'LogDate2')
    log_search_btn = driver.find_element(By.ID, 'Button2')

    log_start_date.send_keys(date_in.strftime("%Y.%m.%d"))
    log_end_date.send_keys(date_out.strftime("%Y.%m.%d"))

    log_search_btn.click()
    sleep(2)
    driver.quit()
    sleep(2)   
    return True        

def parse_file(file_name, events):
    try:
        with open(f"{LOG_PATH}/{file_name}.log", 'r', encoding='utf-8') as file:
            for line in file:
                if line.find('Отпечаток') != -1 or line.find('for') != -1:
                    try:
                        date = re.search(r'\d{2}\.\d{2}\.\d{4}', line)[0]
                    except Exception:
                        date = 'ERROR'
                    try:
                        time = re.search(r'\d{1,2}:\d{2}:\d{2}', line)[0]
                    except Exception:
                        time = 'ERROR'
                    id = re.search(r'ДЮ-\d+', line)
                    if id:
                        id = id[0]
                    if date in log:
                        log[date].append({
                            'time': time,
                            'id': id
                        })
                    else:
                        log[date] = []
                        log[date].append({
                            'time': time,
                            'id': id
                        })
    except Exception:
        print(f'Ошибка доступа к файлу:  "{LOG_PATH}/{file_name}.log"')

def get_timestamp(time, mid_day = MID_DAY_HOUR):
    hour = int(time.split(':')[0])
    if hour < mid_day:
        result = 'Утро'
    else:
        result = 'Вечер'
    return result

def normalize_log():
    result = {}
    for date, events_per_day in log.items():
        result[date] = {}
        for event in events_per_day:
            if event['id'] in result[date]:
                result[date][event['id']].append(event['time'])
            else:
                result[date][event['id']] = []
                result[date][event['id']].append(event['time'])
    return result
  
def set_events():
    events = []
    for group in dpg.get_item_children(events_groups, 1):
        event = {}
        counter = 1
        for input in dpg.get_item_children(group, 1):
            if counter == 1:
                event['worker'] = 'ДЮ-' + dpg.get_value(input)
            if counter == 2:
                event['date'] = datetime.datetime.strptime(dpg.get_value(input), '%d.%m.%Y').date()
            if counter == 3:
                event['is_day'] = dpg.get_value(input)
            counter += 1
        event['pass'] = 0
        if event['is_day'] == 'Day':
            event['confirm'] = []
            event['abort'] = []
        else:
            event['confirm_day1'] = []
            event['confirm_day2'] = []
            event['abort_day1'] = []
            event['abort_day2'] = []

        events.append(event)
    return events

def get_response(event):
    status = event['status']
    date = event['date'].strftime("%d.%m.%Y")
    if 'start_work_date' in event:
        start_work_date = event['start_work_date'].strftime("%d.%m.%Y")
    worker = event['worker']
    passed = event['pass']
    confirmed_mark = event['confirmed_mark']

    SUCC_MARK = f'За {date} по сотруднику {worker} техническая проблема подтверждена. \nДля редактирования табеля следует обратиться в "Табель Учёта Рабочего Времени (Ваш РУ)".\n'
    FAIL_OUT = f'За {date} по сотруднику {worker} в логфайле отсутствуют попытки прикладывания пальца на уход. \nВ связи с этим мы не можем подтвердить техническую проблему.\n'
    FAIL_IN = f'За {date} по сотруднику {worker} в логфайле отсутствуют попытки прикладывания пальца на приход. \nВ связи с этим мы не можем подтвердить техническую проблему.\n'
    FAIL_ALL = f'За {date} по сотруднику {worker} в логфайле отсутствуют попытки прикладывания пальца на приход и уход. \nВ связи с этим мы не можем подтвердить техническую проблему.\n'
    FAIL_PASS_MARK = f'Количество заведенных в систему отпечатков по сотруднику {worker}: {passed}\nПо регламенту их должно быть 2. \nВ связи с этим мы не можем подтвердить техническую проблему.\n'
    SUCC_AXAPTA = f'Передано на отдел сопровождения Axapta. \nЗа {date} по сотруднику {worker} в логфайле зафиксированы следующие события: \n'
    if 'start_work_date' in event:
        FAIL_START_WORK_DATE = f'Дата запрашиваемого события {date}, опережает дату устройства сотрудника {worker} в магазин {start_work_date}\n'
    FAIL_BIO8 = f'Cотрудник не заводил отпечатки в новой версии Biolink v8.\nВ связи с этим мы не можем подтвердить техническую проблему.\nНеобходимо завести отпечатки пальцев по инструкции.\n'
    NOT_FOUND = f'Сотрудник не найден, убедитесь что {worker} корректный'
    pprint(event)

    for mark in confirmed_mark:
        SUCC_AXAPTA += f"{mark['date']} {mark['time']} {mark['about']}\n" 
        
    mess = ''
    if status == 'succ_axapta': mess = SUCC_AXAPTA
    if status == 'succ_mark': mess = SUCC_MARK
    if status == 'fail_out': mess = FAIL_OUT
    if status == 'fail_in': mess = FAIL_IN
    if status == 'fail_all': mess = FAIL_ALL
    if status == 'fail_pass_mark': mess = FAIL_PASS_MARK
    if status == 'fail_start_work_date': mess = FAIL_START_WORK_DATE
    if status == 'fail_bio8': mess = FAIL_BIO8
    if status == 'not_found': mess = NOT_FOUND
    BUFFER.append(mess)
    return mess

def copy():
    res = ''
    for mess in BUFFER:
        res += mess + '\n'
    pyperclip.copy(res)


def create_new_window(events, shop, file_name):
    def open_log_file():
        osCommandString = f"notepad.exe {LOG_PATH}/{file_name}.log"
        os.system(osCommandString)
    
    with dpg.window(label='result', pos=(10, 10), width=760, height=540):
        dpg.add_button(label='COPY TO CLIPBOARD', callback=copy, width=744, height=100)
        dpg.add_button(label='OPEN LOG', callback=open_log_file, width=744, height=50)
        with dpg.table(header_row=True, row_background=True):
                dpg.add_table_column(label="Дата")
                dpg.add_table_column(label="Всего")
                dpg.add_table_column(label="Всего(Отказ)")
                dpg.add_table_column(label="Утро(Отказ)")
                dpg.add_table_column(label="Вечер(Отказ)")

                for date in log:
                    table = TableData(log[date])
                    with dpg.table_row():                      
                        dpg.add_text(default_value=date)
                        dpg.add_text(table.get_count_all_events())
                        dpg.add_text(table.get_count_all_abort())
                        dpg.add_text(table.get_count_morn_abort())
                        dpg.add_text(table.get_count_even_abort())

                for event in events:
                    if event['is_day'] == 'Day':
                        date = event['date'].strftime("%d.%m.%Y")
                        if date in log:
                            pass
                        else:
                            with dpg.table_row():                      
                                dpg.add_text(date, color=RED_COLOR)
                                dpg.add_text(0, color=RED_COLOR)
                                dpg.add_text(0, color=RED_COLOR)
                                dpg.add_text(0, color=RED_COLOR)
                                dpg.add_text(0, color=RED_COLOR)
                    if event['is_day'] == 'Night':
                        date1 = event['date'].strftime("%d.%m.%Y")
                        date2 = event['date'] + datetime.timedelta(days=1)
                        date2 = date2.strftime("%d.%m.%Y")
                        if date1 in log:
                            pass
                        else:
                            with dpg.table_row():                      
                                dpg.add_text(date1, color=RED_COLOR)
                                dpg.add_text(0, color=RED_COLOR)
                                dpg.add_text(0, color=RED_COLOR)
                                dpg.add_text(0, color=RED_COLOR)
                                dpg.add_text(0, color=RED_COLOR)
                        if date2 in log:
                            pass
                        else:
                            with dpg.table_row():                      
                                dpg.add_text(date2, color=RED_COLOR)
                                dpg.add_text(0, color=RED_COLOR)
                                dpg.add_text(0, color=RED_COLOR)
                                dpg.add_text(0, color=RED_COLOR)
                                dpg.add_text(0, color=RED_COLOR)

        dpg.add_separator()
        for event in events:
            color = (0, 255, 0, 255) if event['status'] == 'succ_axapta' or event['status'] == 'succ_mark' else (255, 0, 0, 255)
            dpg.add_text(default_value=f'{shop} по сотруднику {event["worker"]} за {event["date"].strftime("%d.%m.%Y")} ({"Дневная" if event["is_day"] == "Day" else "Ночная"} смена)', color=color)
            with dpg.group(horizontal=True) as confirm_group:
                dpg.add_text('Confirm     :')
                dpg.add_text(f"({len(event['confirmed_mark'])})", color=YELLOW_COLOR)
                with dpg.group(horizontal=False):
                    for mark in event['confirmed_mark']:
                        dpg.add_text(f"({mark['date']} {mark['time']} {mark['about']})")

            if event['is_day'] == 'Day':
                with dpg.group(horizontal=True) as abort_group:
                    dpg.add_text('Abort       :')
                    length = len(event['abort'])
                    dpg.add_text(f"({length})", color=YELLOW_COLOR)
                    if length > 6:
                        dpg.add_text(event['abort'][0])
                        dpg.add_text('...')
                        dpg.add_text(event['abort'][-1])
                    else:
                        for abort in event['abort']:
                            dpg.add_text(abort)

            if event['is_day'] == 'Night':
                with dpg.group(horizontal=True, parent='result') as abort_day1_group:
                    dpg.add_text('Abort DAY 1 :')
                    length = len(event['abort_day1'])
                    dpg.add_text(f"({length})", color=YELLOW_COLOR)
                    if length > 6:
                        dpg.add_text(event['abort_day1'][0])
                        dpg.add_text('...')
                        dpg.add_text(event['abort_day1'][-1])
                    else:
                        for abort in event['abort_day1']:
                            dpg.add_text(abort)
                with dpg.group(horizontal=True, parent='result') as abort_day2_group:
                    dpg.add_text('Abort DAY 2 :')
                    length = len(event['abort_day2'])
                    dpg.add_text(f"({length})", color=YELLOW_COLOR)
                    if length > 6:
                        dpg.add_text(event['abort_day2'][0])
                        dpg.add_text('...')
                        dpg.add_text(event['abort_day2'][-1])
                    else:
                        for abort in event['abort_day2']:
                            dpg.add_text(abort)

            with dpg.group(horizontal=True) as pass_group:
                dpg.add_text('Passed      :')
                passed = event['pass']
                if passed >= 2:
                    text_color = (0, 255, 0, 255)
                else:
                    text_color = (255, 0, 0, 255)
                dpg.add_text(f"({passed})", color=text_color)


                if passed < 2:
                    dpg.add_text('У сотрудника не заведено 2 отпечатка по магазину', color=RED_COLOR)
                    


            with dpg.group(horizontal=True) as status_group:
                dpg.add_text('Status      :')
                dpg.add_text(event['status'])

            with dpg.group(horizontal=True) as response_group:
                dpg.add_text('Response    :')
                dpg.add_text(get_response(event), wrap=600)

            dpg.add_separator()

def add_event(sender, data):
    group = dpg.add_group(horizontal=True, parent=events_groups)
    dpg.add_input_text(label="Worker number", width=100, parent=group)
    dpg.add_input_text(label="Event date", width=100, default_value=DEFAULT_DATE, parent=group)
    dpg.add_radio_button(['Day', 'Night'], horizontal=True, default_value='Day', parent=group)
    ELEMENTS.append(group)

def destroy_elements(sender, data):
    for element in ELEMENTS:
        dpg.delete_item(element)
    ELEMENTS.clear()

def save_log(shop, norm_log):
    for date in norm_log:
        with open(f"./logs/{shop}_{date}.json", 'w') as fh:
            json.dump(norm_log[date], fh) 

def check_file_exsist(file_name):
    if os.path.exists(f"{LOG_PATH}/{file_name}.log"):
        os.remove(f"{LOG_PATH}/{file_name}.log")
    
def find(sender, data):
    log.clear()
    BUFFER.clear()
    with dpg.window(label='preloader', pos=(250, 200), width=300, height=200, no_move=True, no_close=True, no_resize=True, no_collapse=True, no_title_bar=True, modal=True) as window:        
        dpg.add_loading_indicator(pos=(120, 50))
        dpg.add_text('Загружаем лог...', pos=(90, 120))
    query = Query(dpg.get_value('input_shop'), set_events())
    FILENAME = f'Bio_FingerLogs_{query.shop}_{query.get_start_date().strftime("%Y.%m.%d")}-{query.get_end_date().strftime("%Y.%m.%d")}'
    check_file_exsist(FILENAME)
    start_chrome(DRIVER_PATH, query.shop, query.events, query.get_start_date(), query.get_end_date())
    parse_file(FILENAME, query.events)
    norm_log = normalize_log()
    resp = Response(query.events, norm_log)
    resp.make()
    resp.calculate()
    # save_log(query.shop, norm_log)
    create_new_window(query.events, query.shop, FILENAME)
    
       
with dpg.window(label="App", tag="main_window"):
    dpg.bind_font("Default font")
    with dpg.group(horizontal=True) as main_group:
        dpg.add_input_text(label="Shop number", tag='input_shop', width=100)
        dpg.add_button(label="ADD EVENT", callback=add_event, width=100)
        dpg.add_button(label="DESTROY", callback=destroy_elements, width=100)
        dpg.add_button(label="FIND", callback=find, width=400)
    dpg.add_separator()
    with dpg.group(horizontal=False) as events_groups:
        with dpg.group(horizontal=True):
            dpg.add_input_text(label="Worker number", tag='input_worker', width=100)
            dpg.add_input_text(label="Event date", tag='input_date', width=100, default_value=DEFAULT_DATE)
            dpg.add_radio_button(['Day', 'Night'], horizontal=True, default_value='Day')

dpg.create_viewport(title='Bio robot', width=800, height=600)
dpg.setup_dearpygui()
dpg.show_viewport()
dpg.set_primary_window("main_window", True)
dpg.start_dearpygui()
dpg.destroy_context()