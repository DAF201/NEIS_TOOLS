import os
import json

CONFIG = {}

SERIAL_REGEX = "158\d{10}"
DN_REGEX = "\d{8}"
WO_REFEX = "\d{12}-\d{1}"
MO_REGEX = ".{3}-.{5}-.{4}-.{3}"
OCR_FXSJ_PGI_REGEX = "PGI"
OCR_FXSJ_PB_REGEX = "PB-\d{5}"
OCR_FXSJ_AMOUNT_REGEX = "GR \d{1,9}"
OCR_DN_REGEX = "Delivery No \d{8}"
OCR_PB_REGEX = "Build Nr PB-\d{5}"
OCR_RECIPIENT_REGEX = r'(?m)^(\d+)\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)\s+([A-Z]+)\s+(\d{8})\s+(CUBE)\s+(\d+)$'
os.system("taskkill /f /im excel.exe")

# force onedrive to sync target table
with open("config.json", "r") as config:
    CONFIG = json.load(config)


def update_config() -> None:
    with open("config.json", "w") as config:
        json.dump(CONFIG, config)
