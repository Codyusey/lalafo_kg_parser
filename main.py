"""
One-thread parser  https://lalafo.kg/bishkek/kvartiry/prodazha-kvartir/
Parsing speed approximately ~ 120 items/s.
OS: Windows 10
"""

import datetime
import json
import os
import sys
import time
from random import randrange

import colorama
import cursor
import pytz
import requests
from bs4 import BeautifulSoup
from colorama import Fore
from openpyxl import Workbook
from openpyxl.styles import Font
from tqdm import tqdm
from winsound import MessageBeep, MB_OK, MB_ICONHAND

colorama.init(autoreset=True)
current_data_time = datetime.datetime.now().strftime("%d_%m_%Y_%H_%M")
current_data = datetime.datetime.now().strftime("%d_%m_%Y")
abspath_workdir = os.path.dirname((os.path.abspath(__file__)))

headers = {'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;'
                     '=0.8,application/signed-exchange;v=b3;q=0.9',
           'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 ('
                         'KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'}

header_items = ["Number", "Advert id", "Advert url", "Main mobile phone", "Messenger phone_1",
                "Messenger phone_2", "Messenger phone_3", "Messenger phone 4", "Advert title", "Price",
                "Currency", "Created time", "Updated_time"]
MESS_LEN = 90
items_dict = {}
pagination = 0
filename_data_json = "loaded_source.json"


def write_data_json(filename, data):
    try:
        with open(filename, 'w', encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
    except Exception:
        print(f'Error write - {filename}')


def load_data_json(filename=filename_data_json):
    try:
        with open(filename, 'r', encoding="utf-8") as f:
            data_json = json.load(f)
            return data_json
    except Exception:
        print(f'Error read "{filename}"')
        write_data_json(filename, {})  # clear cache
        return {}


def get_pag_data(page):
    """ Get pagination for start loop. Get data json"""
    global pagination
    url = f'https://lalafo.kg/bishkek/kvartiry/prodazha-kvartir?sort_by=default&page={str(page)}'
    count = 0
    while count <= 2:
        with requests.Session() as session:
            try:
                response = session.get(url=url, headers=headers)
                time.sleep(randrange(1, 2))
                soup = BeautifulSoup(response.text, 'html.parser')
                data_str_like_json = soup.find('script', id="__NEXT_DATA__", type="application/json").text
                data_dict = json.loads(data_str_like_json)
                if page == 0:
                    pagination = int(data_dict['props']['initialState']['listing']['listingFeed']['data']["_meta"]['pageCount'])
                    return pagination
                else:
                    return data_dict
            except Exception:
                print_ln(f'Error get data {url}', tab_type='', start_ln='\r', end_ln='', color=Fore.RED)
            if count == 2:
                print_ln("Error has occurred. Please restart the script.", color=Fore.RED)
                beep(b_type=MB_ICONHAND)
                sys.exit()
            count += 1


def parser_json():
    """Scraping pages and save to xls file"""
    pre_id = 0
    for page in tqdm(range(1, get_pag_data(0) + 1), desc='Scraping pages', unit='page', ncols=MESS_LEN,
                     bar_format="{l_bar}%s{bar}%s{r_bar}" % (Fore.GREEN, Fore.RESET)):
        data_dict = get_pag_data(page)
        for key in data_dict['props']['initialState']['listing']['listingFeed']['data']['items']:
            items_lists = [''] * len(header_items)
            items_lists[1] = key['id']
            items_lists[2] = 'https://lalafo.kg' + key['url']
            try:
                items_lists[3] = "tel:" + key['mobile']
            except:
                items_lists[3] = ''
            try:
                items_lists[4] = "tel:" + key['user']['business']['features']['contact_phones']['model']['contacts'][0][
                    'phone']
            except:
                items_lists[4] = ''
            try:
                items_lists[5] = "tel:" + key['user']['business']['features']['contact_phones']['model']['contacts'][1][
                    'phone']
            except:
                items_lists[5] = ''
            try:
                items_lists[6] = "tel:" + key['user']['business']['features']['contact_phones']['model']['contacts'][2][
                    'phone']
            except:
                items_lists[6] = ''
            try:
                items_lists[7] = "tel:" + key['user']['business']['features']['contact_phones']['model']['contacts'][3][
                    'phone']
            except:
                items_lists[7] = ''
            items_lists[8] = key['title']
            try:
                items_lists[9] = key['price']
                items_lists[10] = key['currency']
            except:
                items_lists[9] = ''
                items_lists[10] = ''
            city_timezone = pytz.timezone("Asia/Bishkek")  # local_timezone = tzlocal.get_localzone()
            items_lists[11] = datetime.datetime.fromtimestamp(
                key['created_time'], city_timezone).strftime('%Y-%m-%d %H:%M')
            items_lists[12] = datetime.datetime.fromtimestamp(
                key['updated_time'], city_timezone).strftime('%Y-%m-%d %H:%M')
            if pre_id != key['id']:
                pre_id = key['id']
            items_dict[key['id']] = items_lists

    for num, key in enumerate(items_dict):
        items_dict[key][0] = num + 1
    print_ln(f'Quantity of parsed items: ', start_ln='', tab_type='', end_ln='', color=Fore.GREEN)
    print_ln(f'{len(items_dict)}', start_ln='', tab_type='')  # color = Fore.WHITE


def write_items_xlsx(name_xlsx, data_dict_, header_list_):
    """ write data to xlsx file"""
    work_book_ = Workbook()
    work_cell = work_book_.active
    work_cell.column_dimensions['A'].width = len(header_list_[0])
    work_cell.column_dimensions['B'].width = len(header_list_[1]) + 5
    work_cell.column_dimensions['C'].width = len(header_list_[2]) + 7
    work_cell.column_dimensions['D'].width = len(header_list_[3]) + 3
    work_cell.column_dimensions['E'].width = len(header_list_[4]) + 2
    work_cell.column_dimensions['F'].width = len(header_list_[5]) + 2
    work_cell.column_dimensions['G'].width = len(header_list_[6]) + 2
    work_cell.column_dimensions['H'].width = len(header_list_[7]) + 2
    work_cell.column_dimensions['I'].width = len(header_list_[8]) + 12
    work_cell.column_dimensions['J'].width = len(header_list_[9]) + 2
    work_cell.column_dimensions['K'].width = len(header_list_[10])
    work_cell.column_dimensions['L'].width = len(header_list_[11]) + 4
    work_cell.column_dimensions['M'].width = len(header_list_[12]) + 4
    work_cell.title = f'lalafo_kg_{len(data_dict_)}itm_{current_data}'
    if header_list_ != "":
        work_cell.append(header_list_)
    if len(data_dict_) != 0:
        for row in data_dict_.values():
            work_cell.append(row)
        for cell in work_cell["1:1"]:
            cell.font = Font(bold=True)
        work_book_.save(name_xlsx)
    else:
        print_ln('Noting to write. Data is empty', start_ln='', tab_type='*', color=Fore.GREEN)


def print_ln(message, tab_type='…', start_ln='\n', end_ln='\n', color=Fore.WHITE):
    """ Print messages in line"""
    mess_len = MESS_LEN
    if len(tab_type) == 0:
        spc = ""
    else:
        spc = " "
    tab_len = int((mess_len - 2 - len(message)) / 2)
    print(color + start_ln + tab_type * tab_len + spc + message + spc + " " * (
            len(message) % 2) + tab_type * tab_len, end=end_ln)


def beep(times=1, b_type=MB_OK):
    """Sound notifications"""
    for _ in range(times):
        MessageBeep(b_type)
        time.sleep(1)


def main():
    """Run parser"""
    cursor.hide()

    print_ln("lalafo_kg parser's started", color=Fore.GREEN)
    parser_json()
    xlsx_pathname = 'lalafo_kg_' + str(len(items_dict)) + 'items_' + current_data_time + '.xlsx'
    write_items_xlsx(xlsx_pathname, items_dict, header_items)
    print_ln('End', start_ln='', color=Fore.GREEN)
    beep(3)
    os.startfile(abspath_workdir)
    os.startfile(xlsx_pathname)
    cursor.show()


if __name__ == '__main__':
    main()
