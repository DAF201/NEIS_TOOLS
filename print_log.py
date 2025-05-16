from datetime import datetime
import traceback
from gc import collect

SYS_START_TIME = datetime.now()


def get_line():
    return traceback.extract_stack()[-2].lineno


def message(func_name="", msg=""):
    if func_name == "" or msg == "":
        return
    system_operated_time = (datetime.now() - SYS_START_TIME).total_seconds()
    operate_min = int(system_operated_time / 60)
    operate_sec = int(system_operated_time % 60)
    print("[{}:{}] {}:\t{}".format(operate_min, operate_sec, func_name, msg))
    collect()


def alert(func_name, line, msg):
    system_operated_time = (
        datetime.now() - SYS_START_TIME).total_seconds()
    operate_min = int(system_operated_time / 60)
    operate_sec = int(system_operated_time % 60)
    print("[{}:{}] {}:\t{}\t@Line: {}".format(operate_min,
          operate_sec, func_name, msg, line))
    collect()
