#!/usr/bin/python3
#-*- coding:utf-8 -*-


import requests, openpyxl


"""
Programm for parsing from xt-xarid.uz
(c) GRTB
@gr_team_boss
"""


def write_to_excel(filename: str, start: int, arg: list) -> any:
    if len(arg) == 0:
        return 'Arguments is end!'
    else:
        if start == 0:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(['ID', 'Наименование'])
        else:
            wb = openpyxl.load_workbook(filename=f'{filename}.xlsx')
            ws = wb.active
        for item in arg:
            ws.append([item['product_id'], item['product_name']])
        wb.save(f'{filename}.xlsx')
    return 'OK!'


def main() -> None:
    URI = 'https://api.xt-xarid.uz/rpc'
    filename = input('Введите имя файла(без расширения): ')
    user_value = int(input('Введите значения с шагом 51: '))
    for start in range(0, user_value, 51):
        JSON_REQUEST = {
                "id": 1,
                "jsonrpc": "2.0",
                "method": "ref",
                "params": {
                    "ref": "ref_online_shop_products",
                    "op": "read",
                    "offset": start,
                    "limit": 51,
                    "filters": {},
                    }
                }
        r = requests.post(URI, json=JSON_REQUEST)
        assert r.status_code == 200
        print(r.status_code)
        response = r.json()['result']
        status_response_excel = write_to_excel(filename, start, response)
        if status_response_excel == "OK!":
            print("[*] OK!")
        else:
            print("[*] Proccess is end!")


if __name__ == "__main__":
    main()
