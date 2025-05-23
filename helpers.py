from datetime import datetime
import traceback
from gc import collect
from winsound import Beep
SYS_START_TIME = datetime.now()


def alert_beep(msg):
    Beep(1000, 1000)
    print(msg)


def get_line():
    return traceback.extract_stack()[-2].lineno


def get_runtime():
    system_operated_time = (datetime.now() - SYS_START_TIME).total_seconds()
    operate_hour = system_operated_time // 3600
    operate_min = (system_operated_time % 3600) // 60
    operate_sec = system_operated_time % 60
    return int(operate_hour), int(operate_min), int(operate_sec)


def message(func_name="", msg=""):
    if func_name == "" or msg == "":
        return
    operate_hour, operate_min, operate_sec = get_runtime()
    print("[{}:{}:{}] {}:\t{}".format(str(operate_hour).rjust(2, "0"), str(operate_min).rjust(2, "0"), str(
        operate_sec).rjust(2, "0"), func_name, msg))
    collect()


def alert(func_name, line, msg):
    operate_hour, operate_min, operate_sec = get_runtime()
    print("[{}:{}:{}] {}:\t{}\t@Line: {}".format(str(operate_hour).rjust(2, "0"), str(operate_min).rjust(2, "0"),
          str(operate_sec).rjust(2, "0"), func_name, msg, line))
    collect()
