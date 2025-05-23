import os
import json

CONFIG = {}

SERIAL_REGEX = "158\d{10}"
DN_REGEX = "\d{8}"
WO_REFEX = "\d{12}-\d{1}"
MO_REGEX = ".{3}-.{5}-.{4}-.{3}"

os.system("taskkill /f /im excel.exe")

# force onedrive to sync target table
with open("config.json", "r") as config:
    CONFIG = json.load(config)


def update_config() -> None:
    with open("config.json", "w") as config:
        json.dump(CONFIG, config)
