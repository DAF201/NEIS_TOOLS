from helpers import message
from time import sleep
import os
import win32com.client
from config import CONFIG
from datetime import datetime

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
    with open(os.path.dirname(CONFIG["Excel"]["target_table"])+"sync.txt", "w") as sync:
        sync.write("sync start")
    sleep(1)
    os.remove(os.path.dirname(CONFIG["Excel"]["target_table"])+"sync.txt")
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


def create_report():
    """create the report, based on if there is a DN and the status is SHIPPED"""
    report_init()
    today_date = f"{datetime.today().month}/{datetime.today().day}"

    # search for the first appearance of today's transaction
    first_cell = TARGET_TABLE_SHEET.Range("A:A").Find(
        What=today_date,
        LookIn=1,       # xlValues
        LookAt=1,       # xlWhole (exact match)
        SearchOrder=1,  # xlByRows
        SearchDirection=1  # xlNext (top to bottom)
    )
    start_row = first_cell.Row

    # copy the custom value for date in target table
    excel_today_value = TARGET_TABLE_SHEET.Cells(start_row, 1).Value

    last_row = REPORT_TABLE_SHEET.UsedRange.Rows.Count
    if last_row > 1:
        REPORT_TABLE_SHEET.Range(f"2:{last_row}").ClearContents()
        REPORT_TABLE_WORKBOOK.Save()

    report_row_index = 2
    current_row = start_row

    # for each row, if the value of the DN is FXSJ or something, and the color is green, then we consider it is good for report
    while (TARGET_TABLE_SHEET.Cells(current_row, 2).Value != None):
        if report_check(current_row):

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

            # leave blank for reqeust number and recipient
            REPORT_TABLE_SHEET.Cells(
                report_row_index, 5).Value = ""
            REPORT_TABLE_SHEET.Cells(
                report_row_index, 6).Value = "NAME"

            # copy the GR value
            REPORT_TABLE_SHEET.Cells(
                report_row_index, 7).Value = TARGET_TABLE_SHEET.Cells(current_row, 13).Value

            # copy DN
            REPORT_TABLE_SHEET.Cells(
                report_row_index, 8).Value = TARGET_TABLE_SHEET.Cells(current_row, 12).Value

            # Vendor Pooled? PGI?
            if TARGET_TABLE_SHEET.Cells(current_row, 12).Value == "FXSJ":
                REPORT_TABLE_SHEET.Cells(
                    report_row_index, 9).Value = "Yes"
                REPORT_TABLE_SHEET.Cells(
                    report_row_index, 10).Value = "FXSJ STATUS"
            else:
                REPORT_TABLE_SHEET.Cells(
                    report_row_index, 9).Value = "NA"
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
                REPORT_TABLE_SHEET.Cells(
                    report_row_index, 12).Value = "TRACKING NUMBER"

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
        current_row += 1
    try:
        REPORT_TABLE_WORKBOOK.Save()
        REPORT_TABLE.Quit()
    except:
        pass
    try:
        TARGET_TABLE_WORKBOOK.Save()
        TARGET_TABLE.Quit()
    except:
        pass

    message(__name__, "REPORT OF {} HAS BEEN GENERATED AT {}".format(
        today_date, CONFIG["Excel"]["report_table"]))
