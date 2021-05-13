import datetime
import os
import time
import pandas as pd
from selenium import webdriver
import random
from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.utils import formatdate
from transliterate import translit
import shutil


def create_subdirs(file_names=[]):  # для каждого экселя в Input data создает папку в Output data
    # с тем же именем без расширения + текущая дата
    clear_names = []
    for name in file_names:
        new_name = name.strip('.csv').strip('.xlsx').strip('.xls')
        clear_names.append(new_name)
    date = datetime.datetime.now().strftime('%Y-%m-%d')
    for name in clear_names:
        os.mkdir(f'Output data/{date} {name}')


def get_info_about_input_data(input_filename):
    phys_INNs = []
    company_INNs = []
    df = pd.read_excel(f'Input data/{input_filename}')
    INN_list = df['ИНН'].tolist()
    for value in INN_list:
        if str(value).strip(' ').isdigit() and len(str(value).strip(' ')) == 10:
            company_INNs.append(str(value).strip(' '))
        elif str(value).strip(' ').isdigit() and len(str(value).strip(' ')) == 12:
            phys_INNs.append(str(value).strip(' '))
    return {
        'rows_read': len(INN_list),
        'phys_INNs_count': len(phys_INNs),
        'company_INNs_count': len(company_INNs),
        'file_name': input_filename,
        'file_name_clear': input_filename.strip('.csv').strip('.xlsx').strip('.xls')
    }


def get_input_INNs(input_filename):
    phys_INNs = []
    company_INNs = []
    df = pd.read_excel(f'Input data/{input_filename}')
    INN_list = df['ИНН'].tolist()
    for value in INN_list:
        if str(value).strip(' ').isdigit() and len(str(value).strip(' ')) == 10:
            company_INNs.append(str(value).strip(' '))
        elif str(value).strip(' ').isdigit() and len(str(value).strip(' ')) == 12:
            phys_INNs.append(str(value).strip(' '))
    return company_INNs + phys_INNs


def is_real_INN(INN):
    driver_path = r'C:\Users\Public\MyFirstRPA_Project\FireFoxDriverSelenium\geckodriver.exe'
    driver = webdriver.Firefox(executable_path=driver_path)
    driver.get('https://pb.nalog.ru/search.html#')
    time.sleep(random.randrange(2, 4))
    search_input = driver.find_element_by_id('queryAll')
    search_input.send_keys(f'{INN}')
    time.sleep(random.randrange(2, 4))
    driver.find_element_by_id('quickSubmit').click()
    time.sleep(random.randrange(2, 4))
    result = driver.find_elements_by_class_name('result-group-name')
    time.sleep(random.randrange(2, 4))
    driver.close()
    driver.quit()
    if len(result) > 0:
        return True
    else:
        return False


def get_address_and_ogrn_from_line(line, INN):
    info_list = line.text.split(':')
    if len(str(INN)) == 10:
        address = info_list[0].strip(', ОГРН')
        ogrn = info_list[1].strip(', Дата присвоения ОГРН').strip(' ')
    elif len(str(INN)) == 12:
        address = 'Адрес ИП неизвестен'
        ogrn = info_list[1].strip(', ИНН').strip(' ')
    return address, ogrn


def get_info(INN):  # возвращает словарь с названием/фио, адресом, огрн налогоплательщика
    driver_path = r'C:\Users\Public\MyFirstRPA_Project\FireFoxDriverSelenium\geckodriver.exe'
    driver = webdriver.Firefox(executable_path=driver_path)
    driver.get('https://egrul.nalog.ru/index.html')
    time.sleep(random.randrange(3, 5))
    search_input = driver.find_element_by_id('query')
    search_input.send_keys(f'{INN}')
    time.sleep(random.randrange(3, 5))
    driver.find_element_by_id('btnSearch').click()
    time.sleep(random.randrange(3, 5))
    name = driver.find_element_by_class_name('res-caption').text
    info = driver.find_element_by_class_name('res-text')
    address, ogrn = get_address_and_ogrn_from_line(info, INN)
    output_dict = {'name': name, 'address': address, 'ogrn': ogrn}
    time.sleep(random.randrange(3, 5))
    driver.close()
    driver.quit()
    return output_dict


def create_pdf_per_one_INN(INN, excel_filename, ogrn, name, address):  # создаёт пдф с названием/фио, ОГРН, адресом
    excel_filename_clear = excel_filename.strip('.xls').strip('.xlsx').strip('.csv')
    date = datetime.datetime.now().strftime('%Y-%m-%d')
    canvas = Canvas(f'Output data/{date} {excel_filename_clear}/ИНН {INN}.pdf', pagesize=A4)
    pdfmetrics.registerFont(TTFont('FreeSans', 'Fonts/FreeSans.ttf'))
    canvas.setFont('FreeSans', 12)
    vertical_space = 30
    max_symbols_in_line = 63
    canvas.drawString(30, 800, f'Название/ФИО налогоплательщика:')
    name_chunks = [name[i:i + max_symbols_in_line] for i in range(0, len(name), max_symbols_in_line)]
    i = 1
    for chunk in name_chunks:
        canvas.drawString(60, 800-i*vertical_space, f'{chunk}')
        i = i + 1
    canvas.drawString(30, 750 - vertical_space*len(name_chunks), f'ОГРН: {ogrn}')
    address_chunks = [address[i:i + max_symbols_in_line] for i in range(0, len(address), max_symbols_in_line)]
    canvas.drawString(30, 750 - vertical_space*(len(name_chunks)+1), f'Адрес:')
    i = 1
    for chunk in address_chunks:
        canvas.drawString(60, 750 - vertical_space*(len(name_chunks)+1+i), f'{chunk}')
        i = i + 1
    canvas.showPage()
    canvas.save()


def send_email_report_per_file(rows_read, phys_INNs_read, company_INNs_read,
                               real_phys_INNs, real_company_INNs, filename):
    server = 'smtp.gmail.com'
    user = 'danilo.kiril.auto@gmail.com'
    password = 'MyPassword'

    recipient = 'danilo-kiril@yandex.ru'
    sender = user
    subject = 'Автоматический отчёт по обработанному эксель-файлу с ИНН'
    text = f'Обработан файл \"{filename}\". Прочтено {rows_read} строк\n' + \
                 f'В нём распознано {phys_INNs_read + company_INNs_read} ИНН (без проверки на действительность).\n' + \
                 f'Среди них {phys_INNs_read} ИНН физических лиц и {company_INNs_read} компаний.\n' + \
                 f'Среди них действительных {real_phys_INNs + real_company_INNs} ИНН ' \
                 f'({real_phys_INNs} ИП и {real_company_INNs} компаний).\n\n' + \
                 f'Проверялись ИНН на действительность на сайте pb.nalog.ru. По каждому действительному ИНН ' \
                 f'собрана информация в пдф (фио/назв, адрес, ОГРН). Информация по ИНН бралась с egrul.nalog.ru\n'

    filepath = f'Input data/{filename}'
    translitted_filename = translit(filename, language_code='ru', reversed=True)

    msg = MIMEMultipart()
    msg['From'] = sender
    msg['To'] = recipient
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject
    msg.attach(MIMEText(text))

    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(filepath, "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{translitted_filename}"')
    msg.attach(part)

    # context = ssl.SSLContext(ssl.PROTOCOL_SSLv3)
    # SSL connection only working on Python 3+
    smtp = smtplib.SMTP(server)
    isTls = True
    if isTls:
        smtp.starttls()
    smtp.login(user, password)
    smtp.sendmail(sender, recipient, msg.as_string())
    smtp.quit()


def store_file(filename):  # Убирает обработанный эксель файл в специальную папку. Input data только для непрочитанных
    shutil.move(f'Input data/{filename}', f'Stored input files/{filename}')


if __name__ == '__main__':
    input_filenames = os.listdir('Input data')
    create_subdirs(input_filenames)
    for input_filename in input_filenames:
        input_info = get_info_about_input_data(input_filename)
        input_INNs = get_input_INNs(input_filename)
        real_phys_INNs_count = 0
        real_company_INNs_count = 0
        for INN in input_INNs:
            if is_real_INN(INN):
                if len(str(INN)) == 12:
                    real_phys_INNs_count += 1
                elif len(str(INN)) == 10:
                    real_company_INNs_count += 1
                info = get_info(INN)
                create_pdf_per_one_INN(INN, input_filename, info['ogrn'], info['name'], info['address'])
        send_email_report_per_file(input_info['rows_read'], input_info['phys_INNs_count'],
                                       input_info['company_INNs_count'], real_company_INNs_count,
                                       real_company_INNs_count, input_filename)
        store_file(input_filename)
