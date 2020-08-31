import time
from pprint import pprint

from wialon import flags

def api_wialon_dwnObj(wialon):

    spec1 = {
        'itemsType': 'avl_resource',
        'propName': 'sys_name',
        'propValueMask': '*',
        'sortType': 'sys_name'
    }

    spec2 = {
        'itemsType': 'avl_unit',
        'propName': 'sys_name',
        'propValueMask': '*',
        'sortType': 'sys_name'
    }

    interval = {"from": 0, "to": 0}

    custom_flag = 0x00000001
    units = wialon.core_search_items(spec=spec1, force=1, flags=custom_flag, **interval)
    units2 = wialon.core_search_items(spec=spec2, force=1, flags=custom_flag, **interval)
    items = units['items']
    objs = units2['items']
    total = units2['totalItemsCount']
    resource_id = None

    for res in items:
        if '_api' in res['nm']:
            resource_id = res['id']
            break

    return resource_id, objs, total

