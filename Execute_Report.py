import datetime
import time


def time_conv(time_f, time_zone=7):
    try:

        date, time_d = time_f.split(' ')
        y, m, d = date.split('-')
        h, min, s = time_d.split(':')
        t1 = datetime.datetime(int(y), int(m), int(d), int(h), int(min), int(s))
        time_unix = int(str(time.mktime(t1.timetuple()))[:-2])
        time_unix += time_zone * 3600
        time_f = str(datetime.datetime.fromtimestamp(time_unix))[:-3]
        return time_f

    except:

        y, m, d = time_f.split('-')
        t1 = datetime.datetime(int(y), int(m), int(d))
        time_unix = int(str(time.mktime(t1.timetuple()))[:-2])
        time_unix += time_zone * 3600
        time_f = str(datetime.datetime.fromtimestamp(time_unix))[:11]
        return time_f


def execute_report(wialon, Obj, res, from_time, to_time):
    units = wialon.report_exec_report({
        'reportResourceId': res,
        'reportTemplateId': 1,
        'reportObjectId': Obj.wialon_id,
        'reportObjectSecId': 0,
        'interval': {'from': from_time, 'to': to_time, 'flags': 0}})

    Obj.count_travel = 0

    try:
        Obj.count_travel = int(units['reportResult']['tables'][0]['rows'])
    except:
        pass

    if Obj.count_travel > 0:
        units1 = wialon.report_get_result_rows({
            "tableIndex": 0,
            "indexFrom": 0,
            "indexTo": Obj.count_travel
        })
        return units1
