import json
import requests
from requests.exceptions import RequestException
import re
from xlsxwriter import Workbook
import openpyxl
import pandas as pd

province = ["河北","山西","辽宁","吉林","黑龙","江苏","浙江","安徽","福建","江西",
           "山东","河南","湖北","湖南","广东","海南","四川","贵州","云南","陕西","甘肃",
           "青海","台湾","内蒙","广西","西藏","宁夏","新疆",
           "北京","天津","上海","重庆"]


def get(url):
    try:
        response = requests.get(url)
        if response.status_code == 200:
            return response
        return None
    except RequestException:
        return None


def parse(html):
    data = html.json()
    stations = data['data']

    for item in stations:
        yield {
            'station_name': item['name'],
            'station_address': item['address'],
            'station_province': item['address'][0] + item['address'][1],
            'station_url': item['link']
        }



def write_to_excel(content):
    ordered_list = ["station_name", "station_address", "station_province", "station_url"]
    wb = Workbook("Results.xlsx")
    ws = wb.add_worksheet("Station Data")

    first_row = 0
    for header in ordered_list:
        col = ordered_list.index(header)
        ws.write(first_row, col, header)

    row = 1
    for dic in content:
        for _key, _value in dic.items():
            col = ordered_list.index(_key)
            ws.write(row, col, _value)
        row += 1


    wb.close()


def analyze():
    wb = openpyxl.load_workbook("Results.xlsx")
    ws1 = wb['Station Data']

    province_col = ws1["C"]
    province_list = [province_col[x].value for x in range(len(province_col))]


    for item in province_list:
        if item not in province:
            province_list.remove(item)

    s = pd.Series(province_list).value_counts()
    print(s)
    with pd.ExcelWriter('Results.xlsx',engine='openpyxl',mode='a') as writer:
        s.to_excel(writer, sheet_name='Station Stats')

def main():
    url = 'https://nearby.360che.com/api/UserApi/GetNiaoSuShopList'
    html = get(url)
    station_data = parse(html)
    total_num = 0
    station_list = []

    for item in station_data:
        station_list.append(item)
        total_num = total_num + 1

    print("Total Number of Urea Stations: " + str(total_num))
    write_to_excel(station_list)
    analyze()

if __name__ == '__main__':
    main()
