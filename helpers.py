from datetime import datetime
import traceback
from gc import collect
from winsound import Beep
from glob import glob
from os import path
import pytesseract
from PIL import Image
from pdf2image import convert_from_path
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


def get_today_date() -> str:
    now = datetime.now()
    return now.strftime("%m%d")


def file_search(directory, key):
    res = []
    for file in glob(path.join(directory, "*.*")):
        if key in file:
            res.append(path.abspath(file))
    return res

pytesseract.pytesseract.tesseract_cmd = r".\Tesseract-OCR\tesseract.exe"
poppler_path = r".\Library\bin"
tesseract_config = "--oem 1 --psm 6 -l eng -c tessedit_char_blacklist={}|\/+()‘."


def ocr_reading(image_path) -> str:
    """read a pod to get infomation about a shipment"""
    try:
        images = []
        res = []
        # convert pdf to image for OCR to work
        if image_path.lower().endswith(".pdf"):
            pdf_pages = convert_from_path(
                image_path, poppler_path=poppler_path)
            # for each page of the PDF
            for i, pil_image in enumerate(pdf_pages):
                images.append(pil_image)
        else:
            images.append(Image.open(image_path))
        # for each page, extract text and add to list
        for image in images:
            text = pytesseract.image_to_string(
                image, config=tesseract_config)
            res.append(text)
    except Exception as e:
        alert(__name__, get_line(), e)
    return res
