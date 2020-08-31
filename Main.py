import datetime
import os
from pprint import pprint
from wialon import Wialon, WialonError
import time
from time_set import time_conv
from Download_Obj import api_wialon_dwnObj
from Excel_client import excel_handler
from Execute_Report import execute_report


class Obj:
    Total_total = 0

    def __init__(self):
        self.name_obj = None
        self.wialon_id = None

        self.count_travel = 0
        self.travel_base = {}


def main(path):
    token = 'cc06cce5395ef07d3e3407ae05e79a9808EC7AC81B47A18DA69B534A43958D265B22FB46'
    error = True
    objs = None
    sheet_list = []
    column1 = []
    column2 = []
    column3 = []
    column4 = []

    wialon = Wialon()

    try:
        login = wialon.token_login(token=token)
    except WialonError as e:
        print("Error while login")
        time.sleep(5)
        return
    wialon.sid = login['eid']
    res, obj_base, total = api_wialon_dwnObj(wialon)
    objs = [Obj() for _ in range(total)]
    for num, obj in enumerate(obj_base):
        objs[num].name_obj = obj['nm']
        objs[num].wialon_id = obj['id']
    while True:
        input_sheet = input('1. Отчет за 1 день. \n2. Отчет за период.\n')
        if input_sheet != '1' and input_sheet != '2':
            print('Некорректный выбор')
        elif input_sheet == "1":

            while error:
                try:
                    date_test = input('Введите дату для отчета в формате DD.MM.YYYY (пример: 24.01.2020) \n')
                    from_time, to_time, date = time_conv(date_test)
                    print(from_time)
                    print(to_time)
                    sheet_list.append(date_test)
                    error = False
                except:
                    continue
            for obj in objs:
                try:
                    units = execute_report(wialon, obj, res, from_time, to_time)
                    print(units)
                except:
                    continue




            # handler(path, sheet_list, date_test)

path = os.getcwd()
if __name__ == '__main__':
    main(path)
