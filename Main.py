import datetime
import os
from pprint import pprint
from wialon import Wialon, WialonError
import time
from time_set import time_conv
from Download_Obj import api_wialon_dwnObj
from Execute_Report import execute_report
from excel import excel_handler


class Obj:
    Total_total = 0

    def __init__(self):
        self.name_obj = None
        self.wialon_id = None
        self.count_travel = 0
        self.travel_base = []


def main(path):
    token = 'cc06cce5395ef07d3e3407ae05e79a9808EC7AC81B47A18DA69B534A43958D265B22FB46'
    error = True

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
                    from_time, to_time = time_conv(date_test)
                    error = False
                except:
                    print('ERROR4')
                    continue
            for obj in objs:
                print(obj.name_obj)
                try:
                    units = execute_report(wialon, obj, res, from_time, to_time)
                    if units is not None:
                        for travel in units:
                            unix_date = travel['t1']
                            real_date = datetime.datetime.fromtimestamp(unix_date)
                            obj.travel_base.append([real_date.strftime('%d.%m.%y'), travel['c'][0], travel['c'][1]])

                except:
                    print('ERROR3')
                    continue

            for obj in  objs:
                if len(obj.travel_base) > 0:
                    excel_handler(path, obj, date_test)
            return

        else:
            while error:
                try:

                    print('[A] Не рекомендуется запрашивать отчет периодом больше месяца')
                    period_sheet = input('Введите период для отчета, через дефис (пример 2.01.2020 - 2.02.2020)\n')
                    first_sheet, end_sheet = period_sheet.split('-')
                    from_time, fake = time_conv(first_sheet)
                    fake, to_time = time_conv(end_sheet)
                    error = False
                except:
                    print('ERROR1')
                    continue

            for obj in objs:
                print(obj.name_obj)
                try:
                    units = execute_report(wialon, obj, res, from_time, to_time)
                    if units is not None:
                        for travel in units:
                            unix_date = travel['t1']
                            real_date = datetime.datetime.fromtimestamp(unix_date)
                            obj.travel_base.append([real_date.strftime('%d.%m.%y'), travel['c'][0], travel['c'][1]])
                except:
                    print('ERROR2')
                    continue

            for obj in  objs:
                if len(obj.travel_base) > 0:
                    excel_handler(path, obj, period_sheet)
            return


path = os.getcwd()
if __name__ == '__main__':
    main(path)
