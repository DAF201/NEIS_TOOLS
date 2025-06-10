from helpers import message, alert, get_line, get_today_date, file_search, ocr_reading
from outlook import get_scanned_pdf
from os import startfile, path, remove
from config import CONFIG, OCR_DN_REGEX, OCR_PB_REGEX, OCR_RECIPIENT_REGEX, OCR_FXSJ_PB_REGEX, OCR_FXSJ_AMOUNT_REGEX, OCR_FXSJ_PGI_REGEX
from re import findall,  search
from glob import glob
import shutil


def make_pod() -> None:
    """read emails, download pdf of pod, then read them and extract data to make file name"""

    message(__name__, "START CLEANING PDF")
    for file in glob(path.join(CONFIG["POD"]["buffer_path"], "*.pdf")):
        remove(file)

    message(__name__, "START DOWNLOADING PDF")
    pod_files = get_scanned_pdf()

    message(__name__, "DOWNLOADING COMPLETE")
    attention_needed = []

    message(__name__, "START PROCESSING PDF")
    for file in pod_files:
        text = ocr_reading(file)
        file_data = {"DN": "", "PB": "",
                     "RECIPIENT": [], "AMOUNT": 0, "PATH": file}

        for page in text:

            # check if this is FXSJ vendor pool
            amount = search(OCR_FXSJ_AMOUNT_REGEX, page)
            pgi = search(OCR_FXSJ_PGI_REGEX, page)
            pb = search(OCR_FXSJ_PB_REGEX, page)
            if amount != None and pgi != None and pb != None:
                file_data["DN"] = "FXSJ"
                file_data["AMOUNT"] = amount.group(0)[3:]
                file_data["RECIPIENT"] = "FXSJ"
                file_data["PB"] = pb.group(0)
            else:
                delivery_number = search(OCR_DN_REGEX, page)
                if delivery_number != None:
                    file_data["DN"] = delivery_number.group(0)[-8:]

                for x in findall(OCR_RECIPIENT_REGEX, page):
                    file_data["RECIPIENT"].append((x[3], x[2], x[5]))
                    file_data["AMOUNT"] += int(x[5])

                pb_number = search(OCR_PB_REGEX, page)
                if pb_number != None:
                    file_data["PB"] = pb_number.group(0)[-8:]

        if file_data["DN"] == "" or file_data["PB"] == "" or file_data["AMOUNT"] == 0 or file_data["RECIPIENT"] == []:
            attention_needed.append(file_data["PATH"])

        else:
            message(__name__, "OCR COMPLETE FOR DN: {}".format(
                file_data["DN"]))

            if file_search(CONFIG["POD"]["save_path"], file_data["DN"]) != [] and file_data["DN"] != "FXSJ":
                message(__name__, "POD EXISTED, SKIPPED")
                pass

            else:
                message(__name__, "CREATEING POD FOR DN: {}".format(
                    file_data["DN"]))

                shutil.move(file, path.join(CONFIG["POD"]["save_path"], "{}_{}_{}x_{}.pdf".format(
                    file_data["DN"], file_data["PB"], file_data["AMOUNT"], get_today_date())))

    for file in glob(path.join(CONFIG["POD"]["buffer_path"], "*.pdf")):
        if file not in attention_needed:
            remove(file)

    if attention_needed != []:
        alert(__name__, get_line(), "FOLLOWING SHIPMENT NEED ATTENTION:")
        startfile(CONFIG["POD"]["buffer_path"])
