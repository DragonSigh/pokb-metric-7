import os, time, random, json, shutil

import pandas as pd

from datetime import date, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from loguru import logger

reports_path = os.path.join(os.path.abspath(os.getcwd()), 'reports', 'from_emias')

options = webdriver.ChromeOptions()
options.add_argument('--disable-blink-features=AutomationControlled')
options.add_argument("--start-maximized")
options.add_argument("--disable-extensions")
options.add_argument("--disable-popup-blocking")
options.add_argument("--headless=new")
options.add_experimental_option("prefs", {
  "download.default_directory": reports_path,
  "download.prompt_for_download": False,
  "download.directory_upgrade": True,
  "safebrowsing.enabled": True
})

service = Service('C:\chromedriver\chromedriver.exe')
browser = webdriver.Chrome(options=options, service=service)
actions = ActionChains(browser)

def retry_with_backoff(retries = 5, backoff_in_seconds = 1):
    def rwb(f):
        def wrapper(*args, **kwargs):
          x = 0
          while True:
            try:
              return f(*args, **kwargs)
            except:
              if x == retries:
                raise
              sleep = (backoff_in_seconds * 2 ** x +
                       random.uniform(0, 1))
              time.sleep(sleep)
              x += 1
        return wrapper
    return rwb

def complex_function(x):
    if isinstance(x, str):
        first_name = x.split(' ')[1]
        second_name = x.split(' ')[2]
        last_name = x.split(' ')[3].replace(',', '')
        return f'{first_name} {second_name} {last_name}'
    else:
        return 0

def get_newest_file(path):
    files = os.listdir(path)
    paths = [os.path.join(path, basename) for basename in files]
    return max(paths, key=os.path.getctime)

def wait_for_document_ready(driver):
    WebDriverWait(driver, 60).until(lambda driver: driver.execute_script('return document.readyState;') == 'complete')

def download_wait(directory, timeout, nfiles=None):
    """
    Wait for downloads to finish with a specified timeout.

    Args
    ----
    directory : str
        The path to the folder where the files will be downloaded.
    timeout : int
        How many seconds to wait until timing out.
    nfiles : int, defaults to None
        If provided, also wait for the expected number of files.
    """
    seconds = 0
    dl_wait = True
    while dl_wait and seconds < timeout:
        time.sleep(1)
        dl_wait = False
        files = os.listdir(directory)
        if nfiles and len(files) != nfiles:
            dl_wait = True
        for fname in files:
            if fname.endswith('.crdownload'):
                dl_wait = True
        seconds += 1
    return seconds

def autorization(login_data: str, password_data: str):
    browser.get('http://main.emias.mosreg.ru/MIS/Klimovsk_CGB/Main/Default')
    login_field = browser.find_element(By.XPATH, '//*[@id="Login"]')
    login_field.send_keys(login_data)
    password_field = browser.find_element(By.XPATH, '//*[@id="Password"]')
    password_field.send_keys(password_data)
    # Запомнить меня
    browser.find_element(By.XPATH, '//*[@id="Remember"]').click()
    browser.find_element(By.XPATH, '//*[@id="loginBtn"]').click()
    WebDriverWait(browser, 20).until(EC.invisibility_of_element((By.XPATH, '//*[@id="loadertext"]')))
    element = browser.find_element(By.XPATH, '/html/body/div[8]/div[3]/div/button/span')
    element.click()
    logger.debug('Авторизация пройдена')

def open_emias_report(begin_date, end_date):
    # 1. Заходим в журнал карт медицинских обследований
    logger.debug(f'Открываю Журнал медицинских обследований')
    element = browser.find_element(By.XPATH, '//*[@id="Portlet_23"]/div[2]/div[2]/a/span')
    element.click()
    WebDriverWait(browser, 10).until(EC.number_of_windows_to_be(2))
    browser.switch_to.window(browser.window_handles[1])
    WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="mat-select-value-1"]')))
    # 2. Выставляем фильтры
    # Статус карты — Закрытая
    element = browser.find_element(By.XPATH, '//*[@id="mat-select-value-1"]')
    element.click()
    element = browser.find_element(By.XPATH, '//*[@id="mat-option-2"]')
    element.click()
    # Причина закрытия — Обследование пройдено
    element = browser.find_element(By.XPATH, '//*[@id="mat-select-value-3"]')
    element.click()
    element = browser.find_element(By.XPATH, '//*[@id="mat-option-6"]')
    element.click()
    # Дата закрытия — учетный период
    element = browser.find_element(By.XPATH, '//*[@id="mat-input-9"]')
    ActionChains(browser).click(element).send_keys(begin_date.strftime('%d.%m.%Y')).perform()
    element = browser.find_element(By.XPATH, '//*[@id="mat-input-10"]')
    ActionChains(browser).click(element).send_keys(end_date.strftime('%d.%m.%Y')).send_keys(Keys.RETURN).perform()

    WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mat-select-4"]')))

    logger.debug('Отчет сформирован в браузере')

def save_report():
    logger.debug(f'Начинается сохранение файла с отчетом в папку: {reports_path}')
    # Создать папку с отчётами, если её нет в системе
    try:
        os.mkdir(reports_path)
    except FileExistsError:
        pass
    # Сохранить в Excel
    element = browser.find_element(By.XPATH, '/html/body/app-root/ng-component/div/main/app-disp/div/app-journal/div/div[4]/button')
    element.click()
    download_wait(reports_path, 60)
    #browser.close()
    #browser.switch_to.window(browser.window_handles[0])
    logger.debug('Сохранение файла с отчетом успешно')

# Функция для сохранения датафрейма в Excel с автоподбором ширины столбца
def save_to_excel(dframe: pd.DataFrame, path, index_arg=False):
    with pd.ExcelWriter(path, mode='w', engine='openpyxl') as writer:
         dframe.to_excel(writer, index=index_arg)
         for column in dframe:
            column_width = max(dframe[column].astype(str).map(len).max(), len(column))
            col_idx = dframe.columns.get_loc(column)
            writer.sheets['Sheet1'].column_dimensions[chr(65+col_idx)].width = column_width + 5

def analyze_7_report():
    # Соединение датафреймов из ЕМИАСа в один
    df_list = []
    with os.scandir(reports_path) as it:
        for entry in it:
            if entry.is_file():
                df_temp = pd.read_excel(entry.path, usecols = 'A, C, D, F, G, H, I, K', header=0)
                df_list.append(df_temp)
    df_emias = pd.concat(df_list)
    # Вид обследования - 404 Диспансеризация и 404 Профилактические медицинские осмотры
    df_emias = df_emias[(df_emias['Вид мед. обследования'] == '404н Диспансеризация') | \
                        (df_emias['Вид мед. обследования'] == '404н Профилактические медицинские осмотры')]
    save_to_excel(df_emias, reports_path + '\\' + 'Журнал карт медицинских обследований.xlsx')

@retry_with_backoff(retries=5)
def start_report_saving():
    shutil.rmtree(reports_path, ignore_errors=True) # Очистить предыдущие результаты
    credentials_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), 'auth-emias.json')
    # С начала недели
    first_date = date.today() - timedelta(days=date.today().weekday()) # начало текущей недели
    last_date = date.today() # сегодня
     # Сегодня
    #first_date = date.today()
    #last_date = date.today()
    # За прошлую неделю
    #first_date = date.today() - timedelta(days=date.today().weekday()) - timedelta(days=7) # начало прошлой недели
    #last_date = first_date + timedelta(days=6) # конец недели
    # Задать даты вручную
    #first_date = datetime.datetime.strptime('24.05.2023', '%d.%m.%Y').date()
    #last_date  = datetime.datetime.strptime('25.05.2023', '%d.%m.%Y').date()
    # Открываем данные для авторизации
    logger.debug(f'Выбран период: с {first_date.strftime("%d.%m.%Y")} по {last_date.strftime("%d.%m.%Y")}')
    f = open(credentials_path, 'r', encoding='utf-8')
    data = json.load(f)
    f.close()
    for _departments in data['departments']:        
        for _units in _departments["units"]:
            logger.debug(f'Начинается авторизация')
            autorization(_units['login'], _units['password'])
    open_emias_report(first_date, last_date)
    save_report()
    analyze_7_report()
    #os.remove(get_newest_file(reports_path))

    logger.debug('Выгрузка Показателя 7 успешно завершена')

start_report_saving()

browser.quit()