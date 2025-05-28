import win32com.client
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from config import SERIAL_REGEX, CONFIG, update_config
from helpers import *
from os import path, rename
from datetime import datetime
import re
Tk().withdraw()
# I know this function"s variables naming suck but this function is too long, and the logic requires many var with similar names


def find_target_col(sheet: win32com.client, key: str) -> int:
    """find the col index of the given title(row1)"""
    return sheet.Rows(1).Find(
        What=key,
        LookIn=-4163,
        LookAt=1,
        SearchOrder=1,
        SearchDirection=1
    ).Column


def build_GR_from_feedfile() -> None:
    """build a empty GR file from the Feedfile"""

    message(__name__, "START BUILDING GR FILE FROM FEEDFILE")

    GR_file_path = askopenfilename(title="SELECT FEEDFILE")

    if GR_file_path == "":
        return

    gr_file = win32com.client.Dispatch("Excel.Application")
    gr_file_workbook = gr_file.Workbooks.Open(GR_file_path)

    # assume a xlsx with 1 sheet is feedfile, as usually the Feedfile has only 1 sheet and GR file has 3 sheets
    if gr_file_workbook.Sheets.Count == 1:

        message(__name__, "FEEDFILE DETECTED, START BUILDING GR FILE FROM FEEDFILE")

        while (True):
            sn_start = input(
                "please enter strating serial number, enter quit to Quit\n")
            if sn_start.lower() == "quit":
                message(__name__, "BUILDING CANCELLED")
                gr_file_workbook.Close()
                gr_file.Quit()
                return

            sn_check_res = re.match(SERIAL_REGEX, sn_start)
            if sn_check_res == None:
                continue
            else:
                sn_start = int(sn_check_res.group(0))
                break

        # create GR copy sheet and SN STT sheet
        gr_file_workbook.Sheets.Add(
            Before=gr_file_workbook.Sheets(1)).Name = "Sheet 2"
        gr_file_workbook.Sheets.Add(
            After=gr_file_workbook.Sheets(2)).Name = "Sheet 1"

        # move the original sheet to index 2
        gr_file_workbook.Sheets(3).Move(Before=gr_file_workbook.Sheets(2))

        message(__name__, "COLLECTING WO INFOMATION")

        # search for the size of this WO, using the ACTIVITY_QTY col to find the size rather than fixed index for flexibility
        sn_size = int(gr_file_workbook.Sheets(2).Cells(
            2, find_target_col(gr_file_workbook.Sheets(2), "ACTIVITY_QTY")).Value)

        message(__name__, "START BUILDING")

        # search for the serial number and PASS/FAIL column index
        sn_col = find_target_col(
            gr_file_workbook.Sheets(2), "SERIAL_NUMBER")
        pass_col = find_target_col(
            gr_file_workbook.Sheets(2), "PASS_FAIL_SCRAP")

        # fill serials, change to blue, and set PASS
        for row_num in range(2, sn_size+2):

            # set the SN to start + displacement
            # row_num-2 since row start from 1 and row 1 is useless
            gr_file_workbook.Sheets(2).Cells(
                row_num, sn_col).Value = sn_start+row_num-2

            # set value to number with no decimal value
            gr_file_workbook.Sheets(2).Cells(
                row_num, sn_col).NumberFormat = "0"

            # set text color to blue
            gr_file_workbook.Sheets(2).Cells(
                row_num, sn_col).Font.Color = 16711680

            # put a PASS at the cordinating cell
            # 99% we are GR pass, and most of the time,
            # even when we have SHIPASIS or SCRAP there will be no more than 10 units
            # so just manually copy and paste
            gr_file_workbook.Sheets(2).Cells(
                row_num, pass_col).Value = "PASS"

        gr_file_workbook.Save()
        gr_file.Quit()
        message(__name__, "GR FILE BUILT")
    else:
        message(__name__, "INPUT FILE CANNOT BE PROCESSED")


def build_GR() -> None:
    """Build a GR file from an existing GR file"""

    message(__name__, "START BUILDING GR FILE")

    GR_file_path = askopenfilename(title="SELECT GR FILE")
    if GR_file_path == "":
        return

    # getting the directory of the GR file, since we need to rename it later
    gr_file_base_dir = GR_file_path.replace(
        path.basename(GR_file_path), "")

    gr_file = win32com.client.Dispatch("Excel.Application")
    gr_file_workbook = gr_file.Workbooks.Open(GR_file_path)

    # for easier access, this function is quite large
    gr_file_sheet1 = gr_file_workbook.Sheets(1)
    gr_file_sheet2 = gr_file_workbook.Sheets(2)
    gr_file_sheet3 = gr_file_workbook.Sheets(3)

    # For unknown reason the PB cannot be searched, IDK what happened but for now only PB is using fixed index
    pb_value = gr_file_sheet2.Cells(2, 1).Value

    # find SN column since I just found the SN column is not in a fixed place
    sn_col = find_target_col(gr_file_sheet2, "SERIAL_NUMBER")
    # may get float like str, so convert to float first then int
    sn_start = int(float(gr_file_sheet2.Cells(2, sn_col).Value))
    # ACTIVITY_QTY is always int like, and -1 since start already take 1 place
    sn_end = sn_start+int(gr_file_sheet2.Cells(2, find_target_col(
        gr_file_workbook.Sheets(2), "ACTIVITY_QTY")).Value)-1
    print(sn_start, sn_end)

    # remove contents of the sheet 3
    gr_file_sheet3.Cells.Clear()

    # put SN and STT at sheet 3 title
    gr_file_sheet3.Cells(1, 1).Value = "SN"
    gr_file_sheet3.Cells(1, 2).Value = "STT"

    message(__name__, "PLEASE START SCANNING SN.ENTER \"quit\" TO QUIT. WHEN FINISH, PRESS ENTER TO CONTINUE")
    # a set can avoid duplicated SN
    sn_set = set()
    while True:
        sn = input()
        try:
            if sn.lower() == "quit":
                message(__name__, "SCANNING STOPPED BY USER")
                gr_file_workbook.Close()
                gr_file.Quit()
                return
            # end of scanning
            if sn == "":
                message(
                    __name__, "{} UNITS, CORRECT? (PRESS ENTER TO CONTINUE, quit to Quit, ANYTHING ELSE TO SCAN MORE)".format(len(sn_set)))
                if input() == "":
                    break
                else:
                    continue
            if int(sn) > sn_end or int(sn) < sn_start:
                alert_beep("SN NOT IN RANGE, PLEASE RESCAN LAST BOARD")
                continue
            sn_set.add(int(sn))
        # if sn cannot be convert to int (small sticker near the SN)
        except:
            alert_beep("SN NOT VALID, PLEASE RESCAN LAST BOARD")
            continue

    message(__name__, "SCANNING COMPLETE, START PROCESSING DATA")

    sn_list = sorted(sn_set)

    # fill sheet 3
    row_num = 2
    for sn in sn_list:
        gr_file_sheet3.Cells(row_num, 1).Value = sn
        gr_file_sheet3.Cells(row_num, 1).NumberFormat = "0"
        gr_file_sheet3.Cells(row_num, 2).Value = "PASS"
        row_num += 1

    # clean sheet1, copy title from sheet2 to 1
    gr_file_sheet1.Cells.Clear()

    sn_column = gr_file_sheet2.Columns(sn_col)
    gr_file_sheet2.Rows(1).Copy()
    gr_file_sheet1.Rows(1).PasteSpecial(Paste=-4163)

    # expand filter for searching
    try:
        gr_file_sheet2.ShowAllData()
    except:
        pass

    # start from row 2
    sheet1_row_counter = 2
    for sn in sn_list:
        res = sn_column.Find(
            What=sn,
            LookIn=-4163,
            LookAt=1,
            SearchOrder=1,
            SearchDirection=1
        )

        # get the row, change SN to red, paste to sheet 1
        gr_file_sheet2.Cells(res.Row, sn_col).Font.Color = 255
        gr_file_sheet2.Rows(res.Row).Copy()
        gr_file_sheet1.Rows(sheet1_row_counter).PasteSpecial(Paste=-4163)
        gr_file_sheet1.Rows(sheet1_row_counter).NumberFormat = "0"
        sheet1_row_counter += 1

    message(__name__, "PROCESSING COMPLETE, FINALIZING GR FILE")

    # find the vlookup column if any
    formula_cell = gr_file_sheet2.Rows(2).Find(
        What="=",            # Most formulas start with "="
        LookIn=-4123,        # xlFormulas
        LookAt=1,            # xlWhole
        SearchOrder=1,       # xlByRows
        SearchDirection=1    # xlNext
    )

    try:
        gr_file_sheet1.Columns(formula_cell.Column).Delete()
    except:
        pass
    now = datetime.now()
    today_date = now.strftime("%m")+now.strftime("%d")
    # file name to save
    final_GR_file_name = r"{}{} {}_{}x_{}.xlsx".format(gr_file_base_dir, today_date, pb_value, str(
        len(sn_list)), CONFIG["GR"]["invoice_header"]+str(CONFIG["GR"]["invoice_record"]).rjust(4, "0"))
    # save data
    gr_file_workbook.Save()
    gr_file_workbook.Close(False)
    gr_file.Quit()
    # save next invoice
    CONFIG["GR"]["invoice_record"] = CONFIG["GR"]["invoice_record"]+1
    update_config()
    # rename file
    rename(GR_file_path, final_GR_file_name)
    message(__name__, "GR FILE BUILDING COMPLETE, FIEL CREATE: {}".format(
        final_GR_file_name))
    # now go to do the GR on NV computer
    # exit(0)
