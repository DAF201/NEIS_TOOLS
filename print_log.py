from datetime import datetime
from time import sleep
SYS_START_TIME = datetime.now()


def message(func_name="", msg=""):
    if func_name == "" or msg == "":
        return
    system_operated_time = (datetime.now() - SYS_START_TIME).total_seconds()
    operate_min = int(system_operated_time / 60)
    operate_sec = int(system_operated_time % 60)
    print("[{}:{}] {}:\t{}".format(operate_min, operate_sec, func_name, msg))
