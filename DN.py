import openai
from json import loads
from config import CONFIG
from outlook import get_unread_mails
import win32com.client
from helpers import message, alert, get_line
from time import sleep
import os

openai.api_key = CONFIG["AI"]["openai_key"]
# incase the DN module is mostly working with the same excel workbooks
TARGET_TABLE = TARGET_TABLE_WORKBOOK = TARGET_TABLE_SHEET = BUFFER_TABLE = BUFFER_TABLE_WORKBOOK = BUFFER_TABLE_SHEET = ""


def excel_init():
    print("SYNCING ONEDRIVE")
    with open(os.path.dirname(CONFIG["Excel"]["target_table"])+"sync.txt", "w") as sync:
        sync.write("sync start")
    sleep(1)
    os.remove(os.path.dirname(CONFIG["Excel"]["target_table"])+"sync.txt")
    print("SYNC COMPLETE")
    global TARGET_TABLE, TARGET_TABLE_SHEET, TARGET_TABLE_WORKBOOK, BUFFER_TABLE, BUFFER_TABLE_SHEET, BUFFER_TABLE_WORKBOOK
    message(__name__, "LOADING EXCEL")
    TARGET_TABLE = win32com.client.Dispatch("Excel.Application")
    try:
        TARGET_TABLE.Visible = False
    except:
        pass
    TARGET_TABLE_WORKBOOK = TARGET_TABLE.Workbooks.Open(
        CONFIG["Excel"]["target_table"])
    TARGET_TABLE_SHEET = TARGET_TABLE_WORKBOOK.Sheets(1)
    BUFFER_TABLE = win32com.client.Dispatch("Excel.Application")
    try:
        BUFFER_TABLE.Visible = False
    except:
        pass
    BUFFER_TABLE_WORKBOOK = BUFFER_TABLE.Workbooks.Open(
        CONFIG["Excel"]["buffer_table"])
    BUFFER_TABLE_SHEET = BUFFER_TABLE_WORKBOOK.Sheets(1)
    message(__name__, "EXCEL LOADED")


def clean_buffer_table():
    try:
        message(__name__, "CLEANING BUFFER TABLE")
        used_range = BUFFER_TABLE_SHEET.UsedRange
        last_row = used_range.Row + used_range.Rows.Count - 1
        last_col = used_range.Column + used_range.Columns.Count - 1
        BUFFER_TABLE_SHEET.Range(
            BUFFER_TABLE_SHEET.Cells(1, 1),
            BUFFER_TABLE_SHEET.Cells(last_row, last_col)
        ).ClearContents()
        BUFFER_TABLE_WORKBOOK.Save()
        message(__name__, "ALL CLEAN")
    except:
        pass


def excel_clean_up():
    message(__name__, "CLEAR UP START")
    try:
        TARGET_TABLE_WORKBOOK.Close(False)
    except:
        pass
    try:
        TARGET_TABLE.Quit()
    except:
        pass
    try:
        BUFFER_TABLE_WORKBOOK.Close(False)
    except:
        pass
    try:
        BUFFER_TABLE.Quit()
    except:
        pass
    message(__name__, "CLEAN UP FINISH")


def get_DN_info(msg) -> dict:
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[
            {"role": "system", "content":
             """
             YOU ARE A HELPFUL AGENT AND YOU WILL HELP READING THE EAMILS.
             IN THE OUTPUT, YOU WILL OUTPUT EXACTLY LIKE \{"DN":********, "DEST":"**", PB:"PB-*****", "NUM":"***", "GR_NUMBER":"5202A*****", "VALID":"*"\}
             YOU WILL NEED TO REACH THROUGH THE EMAIL TO FIND THOSE INFORMATIONS AND FILL IN THE CORRECT PLACE.
             DN STAND FOR DELIVERY NUMBER, WHICH IS A 8 DIGITS NUMBER. SOMETIMES THE DN MAY CONCAT WITH DEST SUCH AS "87013010-SC". ALSO IF THERE IS ANY LEADING 0 IN DN, YOU NEED TO REMOVE THE LEADING 0.
             DEST STAND FOR DESTINATION, WHICH IS THE SHIPPING DESTINGATION, AND FOR CHINA, USE CN, FOR TAIWAN, USE TW, FOR HONGKONG, USE HK, IN THE VALUE.
             PB STAND FOR PB-NUMBER, WHICH IS A IDENTIFIER FOR COMPOENTS BEING SHIPPED, AND IT IS "PB-" CONCATE WITH A 5 DIGITS FOR NUMBER.
             NUM STAND FOR THE NUMBER OF THE COMPUNENTS TO BE SHIPPED, USUALLY REPRESENTED AS A DECIMAL NUMBER CONCAT WITH A CHARACTER "x". WHEN PROCESSING DATA, REMOVE THE "x".
             GR_NUMBER IS A UNIQUE ID FOR A TRANSACTION, AND IT ALWAYS START WITH 5202A THEN CONCAT TO 5 DIGITS OF NUMBERS.
             VALID STAND FOR IF YOU CAN FIND ALL VALUES FROM THE EMAIL. IF THERE IS AT LEAST ONE VALUE YOU CANNOT FIND, FILL WITH NA, AND SET VALID="0". NO MAKE UP DATA IF YOU CANNOT FIND THE COORDINATE DATA FROM EMAIL.
             IF YOU CAN FIND ALL VALUES FROM THE EMAIL, VALID="1".
             PLASE NOTE, FOR DESTINATION, THE ZANKER AND SC ARE THE SAME, YOU SHOULD USE SC FOR BOTH CASE.
             ALSO, WHEN YOU CANNOT FIND DN, OR TOLD NO DN, IF THE EMAIL SAID FXSJ, VENDER POOL, STOCK, THIS MEANS THE DN IS "FXSJ" AND DEST IS ALSO "FXSJ", AND THIS IS STILL VALID=1 CASE.
             IN ALL CASES, YOUR RETURN VALUES SHOULD BE STRING, AND THE VALUE SHOULD BE ROUNDED BY " TO MAKE JSON LOADABLE.
             """},
            {"role": "user", "content": msg}
        ],
        temperature=0
    )
    try:
        return loads(response["choices"][0]["message"]["content"])
    except:
        return {}


def check_new_DN() -> list:
    clean_buffer_table()
    message(__name__, "READING DN")
    emails = get_unread_mails()
    dn_info = []
    valid_dn = []
    for email in emails:
        dn_info.append(get_DN_info(str(email)))
    for dn in dn_info:
        try:
            if dn["VALID"] == "1":
                message(__name__, dn)
                valid_dn.append(dn)
        except:
            continue
    message(__name__, "READING COMPLETE")
    return valid_dn


def invoice_search(transaction: str | dict) -> int:
    message(__name__, "SEARCHING FOR: {}".format(
        transaction["GR_NUMBER"]))
    GR_invoice = transaction["GR_NUMBER"]
    invoice_col = TARGET_TABLE_SHEET.Columns(22)
    res = invoice_col.Find(
        What=GR_invoice,
        LookIn=-4163,
        LookAt=1,
        SearchOrder=1,
        SearchDirection=1
    )
    if res == None:
        message(__name__, "NO MATCH")
        return 0
    else:
        message(__name__, "MATCH FOUND: {}".format(res.Row))
        return res.Row


def find_next_blank_row(start_row=1, max_columns=26) -> int:
    row = start_row
    while True:
        empty = True
        for col in range(1, max_columns + 1):
            if BUFFER_TABLE_SHEET.Cells(row, col).Value is not None:
                empty = False
                break
        if empty:
            return row
        row += 1


def copy_row(source_row_num: int, target_row_num: int) -> None:
    # copy the row from target table to the buffer table
    max_col = TARGET_TABLE_SHEET.UsedRange.Columns.Count  # Or hardcode if known

    # Manually copy each cell"s value
    for col in range(1, max_col + 1):
        BUFFER_TABLE_SHEET.Cells(target_row_num, col).Value = TARGET_TABLE_SHEET.Cells(
            source_row_num, col).Value
    # Save and clean up
    BUFFER_TABLE_WORKBOOK.Save()


def edit_buffer_table(transaction: str | dict, row_number: int) -> int:
    # this will search for the last appearance of the transaction based on PB number

    # if not dictionary, load as dictionary
    if type(transaction) == str:
        transaction = loads(transaction)
    BUFFER_TABLE_SHEET.Cells(row_number, 12).Value = transaction["DN"]
    BUFFER_TABLE_SHEET.Cells(row_number, 13).Value = transaction["NUM"]
    BUFFER_TABLE_SHEET.Cells(row_number, 14).Value = transaction["DEST"]
    if transaction["DEST"] == "SC":
        BUFFER_TABLE_SHEET.Cells(row_number, 15).Value = "Driver"

    # DN out, move to packing
    BUFFER_TABLE_SHEET.Cells(row_number, 16).Value = "Packing"

    # if not shipping to CN or TW, no pre-alert
    if transaction["DEST"] not in ["CN", "TW"]:
        BUFFER_TABLE_SHEET.Cells(row_number, 17).Value = "NA"

    # if rework cell is not empty, SO DN and SCAN should be NA
    if BUFFER_TABLE_SHEET.Cells(row_number, 4).Value is not None:
        BUFFER_TABLE_SHEET.Cells(row_number, 19).Value = "NA"
        BUFFER_TABLE_SHEET.Cells(row_number, 20).Value = "NA"
        BUFFER_TABLE_SHEET.Cells(row_number, 21).Value = "NA"
    # Save changes
    BUFFER_TABLE_WORKBOOK.Save()


def append_new_DN_to_excel(new_dn: list):
    excel_init()
    message(__name__, "APPENDING DN INFOMATION TO BUFFER TABLE")
    for dn in new_dn:
        try:
            source_row = invoice_search(dn)
            if source_row != 0:
                target_row = find_next_blank_row()
                copy_row(source_row, target_row)
                edit_buffer_table(dn, target_row)
            else:
                message(__name__, "INVOICE NOT FOUND, START PB SEARCH")
                source_row = PB_search(dn)
                if source_row == []:
                    message(__name__, "PB SEARCH NO FOUND")
                for row in source_row:
                    target_row = find_next_blank_row()
                    copy_row(source_row, target_row)
                    edit_buffer_table(dn, target_row)
        except Exception as e:
            alert(__name__, get_line(), e)
            continue
    message(__name__, "DN INFORMATION ADDED ")
    excel_clean_up()


def PB_search(transaction: str | dict) -> list[int]:
    if isinstance(transaction, str):
        transaction = loads(transaction)
    message(__name__, "SEARCHING FOR: {}".format(
        transaction["PB"]))
    PB_value = transaction["PB"]
    NUM_value = int(transaction["NUM"])  # Compare as string
    # Column C, Number of units from email
    PB_col = TARGET_TABLE_SHEET.Columns(3)
    row_nums = []
    first_found = PB_col.Find(
        What=PB_value,
        LookIn=-4163,
        LookAt=1,
        SearchOrder=1,
        SearchDirection=1
    )
    found = first_found
    while found:
        row = found.Row
        num_in_row = int(TARGET_TABLE_SHEET.Cells(
            row, 11).Value)

        if num_in_row == NUM_value:
            message(__name__, "MATCH FOUND: {}".format(row))
            row_nums.append(row)

        found = PB_col.FindNext(found)
        if found is None or found.Address == first_found.Address:
            message(__name__, "NO MATCH")
            break
    message(__name__, "SEARCH COMPLETE")
    return row_nums
