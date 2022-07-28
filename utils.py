import time
from datetime import datetime


def date_today_by_int() -> list[int]:
    return [int(i) for i in str(datetime.date(datetime.now())).split('-')]


def date_to_seconds(date_string):
    if len(date_string.split('.')[-1]) == 2:
        new_date = date_string.split('.')
        new_date[-1] = '20' + new_date[-1]
        date_string = '.'.join(new_date)
    struct_date = time.strptime(date_string, "%d.%m.%Y")
    return time.mktime(struct_date)
