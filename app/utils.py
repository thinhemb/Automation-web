import json
from datetime import datetime, timedelta
import pandas as pd
import copy


def read_json(path, clear=True):
    with open(path, 'r', encoding="utf-8") as file:
        data = json.load(file)
    if clear:
        with open(path, 'w', encoding="utf-8") as file:
            file.write(json.dumps([], indent=4))
    return data


def write_json(path, data):
    with open(path, 'r', encoding="utf-8") as file:
        data_ = copy.deepcopy(list(json.load(file)))
    with open(path, 'w', encoding="utf-8") as file:
        data_.extend(data)
        file.write(json.dumps(data_, indent=4))


def write_file_txt(path, data, mode='w', end=True):
    with open(path, mode) as file:
        for item in data:
            if end:
                item = item + "\n"
            file.write(item)


def read_csv(path):
    df = pd.read_csv(path, encoding='utf-8')
    return df


def read_file_txt(path):
    with open(path, 'r') as file:
        data = file.readlines()
        return data


def count_day(sending_time_start):
    now_day = datetime.now().strftime("%m/%d/%Y")
    end = datetime.strptime(now_day, "%m/%d/%Y")
    start = datetime.strptime(sending_time_start, "%m/%d/%Y")

    delta = end - start
    num_days = delta.days
    return num_days


def set_end_time(start):
    start = datetime.strptime(start, "%H:%M:%S")

    end_time_datetime = start + timedelta(hours=1)
    end_time_one_slot = end_time_datetime.strftime("%H:%M:%S")
    return str(end_time_one_slot)
