
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
