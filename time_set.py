import datetime
import time


def time_conv(date_f):
    try:
        RU_MONTH_VALUES = {

            1 : 'янв.',
            2 : 'фев.',
            3 : 'мар.',
            4 : 'апр',
            5 : 'май',
            6 : 'июн',
            7 : 'июл',
            8 : 'авг.',
            9 : 'сент.',
            10 : 'окт.',
            11 : 'нояб.',
            12 : 'дек'
        }
        d, m, y = date_f.split('.')
        now = datetime.datetime.now()
        if int(d) > 32 or int(m) > 12 or int(y) < 2018 or int(d) < 0 or int(m) < 0 or int(y) > int(now.year):
            print('Неверная дата')
            return
        if int(d) > int(now.day) and int(m) > int(now.month) and int(y) > int(now.year):
            print('Неверная дата')
            return
        date = d + ' ' +RU_MONTH_VALUES[int(m)]
        h1 = '00'
        min1 = '00'
        h2 = '23'
        min2 = '59'
        s = '59'

        t1 = datetime.datetime(int(y), int(m), int(d), int(h1), int(min1), int(s))
        t1_s1_unix = int(str(time.mktime(t1.timetuple()))[:-2])

        t2 = datetime.datetime(int(y), int(m), int(d), int(h2), int(min2), int(s))
        t2_s1_unix = int(str(time.mktime(t2.timetuple()))[:-2])

        return t1_s1_unix, t2_s1_unix

    except Exception as exc:
        print(exc)
        print('Неверный формат даты')


