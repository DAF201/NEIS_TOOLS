from helpers import message, file_search, ocr_reading
from time import sleep
import os
import win32com.client
from config import CONFIG, OCR_FXSJ_AMOUNT_REGEX, OCR_DN_REGEX, OCR_FXSJ_PB_REGEX, OCR_FXSJ_PGI_REGEX, OCR_PB_REGEX, OCR_RECIPIENT_REGEX
from datetime import datetime
from re import search, findall
TARGET_TABLE = TARGET_TABLE_WORKBOOK = TARGET_TABLE_SHEET = ""
REPORT_TABLE = REPORT_TABLE_WORKBOOK = REPORT_TABLE_SHEET = ""


def report_check(report_row_index) -> bool:

    # check if green
    if TARGET_TABLE_SHEET.Cells(report_row_index, 12).Interior.Color == 5296274.0:
        # check if DN has value
        if (TARGET_TABLE.Cells(report_row_index, 12).Value != "Pending") and (TARGET_TABLE.Cells(report_row_index, 12) != None):
            # check if there is FXSO, FXDN, and SFC SCAN
            if (TARGET_TABLE_SHEET.Cells(report_row_index, 19) and TARGET_TABLE_SHEET.Cells(report_row_index, 20) and TARGET_TABLE_SHEET.Cells(report_row_index, 21)):
                return True
    return False


def report_init():
    """because the onedrive target table is not uptodate, so create a file to trigger sync then delete it"""
    print("SYNCING ONEDRIVE")
    with open(os.path.dirname(CONFIG["Excel"]["target_table_folder"])+"sync.txt", "w") as sync:
        sync.write("sync start")
    sleep(5)
    os.remove(os.path.dirname(CONFIG["Excel"]
              ["target_table_folder"])+"sync.txt")
    print("SYNC COMPLETE")

    global TARGET_TABLE, TARGET_TABLE_SHEET, TARGET_TABLE_WORKBOOK, REPORT_TABLE, REPORT_TABLE_WORKBOOK, REPORT_TABLE_SHEET

    message(__name__, "LOADING EXCEL")

    TARGET_TABLE = win32com.client.Dispatch("Excel.Application")

    try:
        TARGET_TABLE.Visible = False
    except:
        pass

    TARGET_TABLE_WORKBOOK = TARGET_TABLE.Workbooks.Open(
        CONFIG["Excel"]["target_table"])

    TARGET_TABLE_SHEET = TARGET_TABLE_WORKBOOK.Sheets(1)

    # report table
    REPORT_TABLE = win32com.client.Dispatch("Excel.Application")

    try:
        REPORT_TABLE.Visible = False
    except:
        pass

    REPORT_TABLE_WORKBOOK = REPORT_TABLE.Workbooks.Open(
        CONFIG["Excel"]["report_table"])
    REPORT_TABLE_SHEET = REPORT_TABLE_WORKBOOK.Sheets(1)


def extract_report_data(text):
    """extrat data from text from OCR"""
    # buffer
    pod_data = {"DN": "", "PB": "", "TRANSACTION": [], "AMOUNT": 0}
    # find the DN PB and RECIPIENT data from text using REGEX (for ordinary POD only)
    dn = search(OCR_DN_REGEX, text)
    if dn != None:
        pod_data["DN"] = dn.group(0)[-8:]
    pb = search(OCR_PB_REGEX, text)
    if pb != None:
        pod_data["PB"] = pb.group(0)[-8:]
    for recipient in findall(OCR_RECIPIENT_REGEX, text):
        pod_data["TRANSACTION"].append(
            (recipient[3], recipient[1], recipient[5]))
        pod_data["AMOUNT"] += int(recipient[5])
    return pod_data


def create_report():
    """create the report, based on if there is a DN and the status is SHIPPED"""

    report_init()
    today_date = f"{datetime.today().month}/{datetime.today().day}"
    # today_date = "6/5"

    message(__name__, "SEARCHING FOR STARTING ROW")

    # search for the first appearance of today's transaction
    first_cell = TARGET_TABLE_SHEET.Range("A:A").Find(
        What=today_date,
        LookIn=1,       # xlValues
        LookAt=1,       # xlWhole (exact match)
        SearchOrder=1,  # xlByRows
        SearchDirection=1  # xlNext (top to bottom)
    )

    start_row = first_cell.Row

    message(__name__, "STARTING ROW FOUND AT :{}".format(start_row))

    # copy the custom value for date in target table
    excel_today_value = TARGET_TABLE_SHEET.Cells(start_row, 1).Value

    message(__name__, "CLEANING REPORT TABLE")

    # clean report table
    last_row = REPORT_TABLE_SHEET.UsedRange.Rows.Count
    if last_row > 1:
        REPORT_TABLE_SHEET.Range(f"2:{last_row}").ClearContents()
        REPORT_TABLE_WORKBOOK.Save()

    message(__name__, "CLEANING COMPLETE, START BUILDING REPORT")

    report_row_index = 2
    current_row = start_row

    attention_needed = []
    # for each row, if the value of the DN is FXSJ or something, and the color is green, then we consider it is good for report
    while (TARGET_TABLE_SHEET.Cells(current_row, 2).Value != None):

        message(__name__, "CURRENTLY PROCESSING ROW: {}".format(current_row))
        if report_check(current_row):

            dn = TARGET_TABLE_SHEET.Cells(current_row, 12).Value
            if dn == "FXSJ":
                message(
                    __name__, "VENDOR POOL FOUND, FILLING DATA FROM TARGET TABLE")

                # write date
                REPORT_TABLE_SHEET.Cells(
                    report_row_index, 1).Value = excel_today_value

                # copy PB
                REPORT_TABLE_SHEET.Cells(
                    report_row_index, 2).Value = TARGET_TABLE_SHEET.Cells(current_row, 3).Value

                # copy WO PN
                REPORT_TABLE_SHEET.Cells(
                    report_row_index, 3).Value = TARGET_TABLE_SHEET.Cells(current_row, 7).Value
                REPORT_TABLE_SHEET.Cells(
                    report_row_index, 4).Value = TARGET_TABLE_SHEET.Cells(current_row, 8).Value

                # request number and recipient
                REPORT_TABLE_SHEET.Cells(
                    report_row_index, 5).Value = "NA"
                REPORT_TABLE_SHEET.Cells(
                    report_row_index, 6).Value = "NA"

                # Qty, copy GR value
                REPORT_TABLE_SHEET.Cells(
                    report_row_index, 7).Value = TARGET_TABLE_SHEET.Cells(current_row, 13).Value

                # DN
                REPORT_TABLE_SHEET.Cells(
                    report_row_index, 8).Value = "NA"

                # Vendor Pooled? yes
                REPORT_TABLE_SHEET.Cells(
                    report_row_index, 9).Value = "Yes"

                # status, depends on if the 18 column has anything? if Yes likely GI to WO, copy over else Vendor Pool
                if TARGET_TABLE_SHEET.Cells(current_row, 18).Value is None:
                    REPORT_TABLE_SHEET.Cells(
                        report_row_index, 10).Value = "FXSJ Pooled Vendor Stock"
                else:
                    REPORT_TABLE_SHEET.Cells(
                        report_row_index, 10).Value = TARGET_TABLE_SHEET.Cells(current_row, 18).Value

                # carrier
                REPORT_TABLE_SHEET.Cells(
                    report_row_index, 11).Value = "NA"

                # tracking
                REPORT_TABLE_SHEET.Cells(
                    report_row_index, 12).Value = "NA"

                # PGI
                REPORT_TABLE_SHEET.Cells(
                    report_row_index, 13).Value = "Done"

                # SAP SO DN
                REPORT_TABLE_SHEET.Cells(
                    report_row_index, 14).Value = TARGET_TABLE_SHEET.Cells(current_row, 19).Value
                REPORT_TABLE_SHEET.Cells(
                    report_row_index, 15).Value = TARGET_TABLE_SHEET.Cells(current_row, 20).Value

                # SFC SCAN
                REPORT_TABLE_SHEET.Cells(
                    report_row_index, 16).Value = TARGET_TABLE_SHEET.Cells(current_row, 21).Value

                # FX GR invoice
                REPORT_TABLE_SHEET.Cells(
                    report_row_index, 17).Value = TARGET_TABLE_SHEET.Cells(current_row, 22).Value

                # POD
                REPORT_TABLE_SHEET.Cells(
                    report_row_index, 18).Value = "Done"

                REPORT_TABLE_SHEET.Cells(report_row_index, 19).Value = "NA"
                report_row_index += 1
            else:
                # for verifying the total amount to avoid the OCR missed part of a transaction
                total_amount = TARGET_TABLE_SHEET.Cells(current_row, 13).Value
                try:
                    message(__name__, "SEARCHING FOR POD")
                    # try to find POD, if DN is something like "MPL" will be added to attention list (already knows not FXSJ)
                    pod_path = file_search(
                        CONFIG["POD"]["save_path"], int(dn))[0]

                    # for make up a complete data of a transaction, because the data may lay on different pages
                    transaction_data = {"DN": "", "PB": "",
                                        "TRANSCATION": [], "AMOUNT": 0}

                    message(__name__, "READING POD OF DN: {}".format(dn))
                    # find the POD, scan each page and extract data
                    for page in ocr_reading(pod_path):
                        # extract data from each page and try to make up the complete data
                        page_data = extract_report_data(page)
                        if page_data["DN"] != "":
                            transaction_data["DN"] = page_data["DN"]
                        if page_data["PB"] != "":
                            transaction_data["PB"] = page_data["PB"]
                        if page_data["TRANSACTION"] != []:
                            transaction_data["TRANSCATION"] = page_data["TRANSACTION"]
                        if page_data["AMOUNT"] != 0:
                            transaction_data["AMOUNT"] = page_data["AMOUNT"]

                    if transaction_data["DN"] != "" and transaction_data["PB"] != "" and transaction_data["AMOUNT"] != 0 and transaction_data["TRANSCATION"] == []:
                        message(__name__, "OCR DATA INCOMPLETE")
                        raise Exception("OCR CANNOT FETCH TRANSACTION")
                    else:
                        if total_amount != transaction_data["AMOUNT"]:
                            message(
                                __name__, "AMOUNT MISMATCH")
                            raise Exception("AMOUNT NOT MATCH")
                        for transaction in transaction_data["TRANSCATION"]:
                            message(__name__, "ADDING REQUEST:{} TO REPORT".format(
                                transaction[0]))
                            # put today value
                            REPORT_TABLE_SHEET.Cells(
                                report_row_index, 1).Value = excel_today_value

                            # copy the PB
                            REPORT_TABLE_SHEET.Cells(
                                report_row_index, 2).Value = TARGET_TABLE_SHEET.Cells(current_row, 3).Value

                            # copy WO, PN
                            REPORT_TABLE_SHEET.Cells(
                                report_row_index, 3).Value = TARGET_TABLE_SHEET.Cells(current_row, 7).Value
                            REPORT_TABLE_SHEET.Cells(
                                report_row_index, 4).Value = TARGET_TABLE_SHEET.Cells(current_row, 8).Value

                            # reqeust number, name, and qty
                            REPORT_TABLE_SHEET.Cells(
                                report_row_index, 5).Value = transaction[0]
                            REPORT_TABLE_SHEET.Cells(
                                report_row_index, 6).Value = transaction[1]
                            REPORT_TABLE_SHEET.Cells(
                                report_row_index, 7).Value = transaction[2]

                            # copy DN
                            REPORT_TABLE_SHEET.Cells(
                                report_row_index, 8).Value = TARGET_TABLE_SHEET.Cells(current_row, 12).Value

                            # vender pool
                            REPORT_TABLE_SHEET.Cells(
                                report_row_index, 9).Value = "NA"

                            # statu
                            REPORT_TABLE_SHEET.Cells(
                                report_row_index, 10).Value = "Shipped"

                            # carrier
                            REPORT_TABLE_SHEET.Cells(
                                report_row_index, 11).Value = TARGET_TABLE_SHEET.Cells(current_row, 15).Value

                            # tracking
                            if TARGET_TABLE_SHEET.Cells(current_row, 15).Value != "Driver" or TARGET_TABLE_SHEET.Cells(current_row, 15).Value != "NA":
                                REPORT_TABLE_SHEET.Cells(
                                    report_row_index, 12).Value = TARGET_TABLE_SHEET.Cells(current_row, 15).Value
                            else:
                                message(
                                    __name__, "INTERNATIONAL SHIPMENT FOUND, TRACKING NUMBER REQUIRED")
                                REPORT_TABLE_SHEET.Cells(
                                    report_row_index, 12).Value = "TRACKING NUMBER REQUIRED"
                                REPORT_TABLE_SHEET.Cells(
                                    report_row_index, 12).Interior.Color == 255.0

                            # PGI
                            REPORT_TABLE_SHEET.Cells(
                                report_row_index, 13).Value = "Done"
                            # SAP SO
                            REPORT_TABLE_SHEET.Cells(
                                report_row_index, 14).Value = TARGET_TABLE_SHEET.Cells(current_row, 19).Value
                            # SAP DN
                            REPORT_TABLE_SHEET.Cells(
                                report_row_index, 15).Value = TARGET_TABLE_SHEET.Cells(current_row, 20).Value
                            # SFC SCAN
                            REPORT_TABLE_SHEET.Cells(
                                report_row_index, 16).Value = TARGET_TABLE_SHEET.Cells(current_row, 21).Value
                            # FX GR invoice
                            REPORT_TABLE_SHEET.Cells(
                                report_row_index, 17).Value = TARGET_TABLE_SHEET.Cells(current_row, 22).Value
                            # POD
                            REPORT_TABLE_SHEET.Cells(
                                report_row_index, 18).Value = "Done"
                            # Pre Alert
                            if TARGET_TABLE_SHEET.Cells(
                                    report_row_index, 14).Value == "CN" or TARGET_TABLE_SHEET.Cells(
                                    report_row_index, 14).Value == "TW":
                                REPORT_TABLE_SHEET.Cells(
                                    report_row_index, 19).Value = "Done"
                            else:
                                REPORT_TABLE_SHEET.Cells(
                                    report_row_index, 19).Value = "NA"
                            report_row_index += 1
                except:
                    attention_needed.append(current_row)
            REPORT_TABLE_WORKBOOK.Save()
        current_row += 1
    message(__name__, "REPORT GENERATING COMPLETE")
    message(__name__, "FOLLOWING ROWS NEED TO BE HANDLED MANUALLY")
    for row in attention_needed:
        report = str(row)+"\tF"
        report += str(TARGET_TABLE_SHEET.Cells(row, 3).Value)+"\t"
        report += str(TARGET_TABLE_SHEET.Cells(row,
                      12).Value).replace(".0", "")+"\t"
        report += str(int(TARGET_TABLE_SHEET.Cells(row, 9).Value))+"\t"
        report += str(TARGET_TABLE_SHEET.Cells(row, 22).Value)+"\t"
        report += str(TARGET_TABLE_SHEET.Cells(row, 23).Value) + "\t"
        print(report)
    REPORT_TABLE_WORKBOOK.Close()
