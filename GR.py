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


def build_GR_from_feedfile():
    message(__name__, "START BUILDING GR FILE FROM FEEDFILE")
    GR_file_path = askopenfilename(title="Select file to attach")
    if GR_file_path == "":
        return
    gr_file = win32com.client.Dispatch("Excel.Application")
    gr_file_workbook = gr_file.Workbooks.Open(GR_file_path)
    # if this file is a feed file, then need to create other 2 sheets
    if gr_file_workbook.Sheets.Count == 1:
        message(__name__, "FEEDFILE DETECTED, START BUILDING GR FILE FROM FEEDFILE")
        # create gr copy sheet and SN STT sheet
        gr_file_workbook.Sheets.Add(
            Before=gr_file_workbook.Sheets(1)).Name = "Sheet 2"
        gr_file_workbook.Sheets.Add(
            After=gr_file_workbook.Sheets(2)).Name = "Sheet 1"
        # move the original sheet to index 2
        gr_file_workbook.Sheets(3).Move(Before=gr_file_workbook.Sheets(2))
        sn_size = int(gr_file_workbook.Sheets(2).Cells(2, 4))
        message(__name__, "COLLECTING INFOMATION")
        # get SN start until get a valid sn or quit command
        while (True):
            sn_start = input(
                "please enter strating serial number, enter quit to Quit\n")
            if sn_start.lower() == "quit":
                return
            # check if the starting SN is valid
            sn_check_res = re.match(SERIAL_REGEX, sn_start)
            # not valid, continue
            if sn_check_res == None:
                continue
            # valid, move to next step
            else:
                sn_start = int(sn_check_res.group(0))
                break
        message(__name__, "START BUILDING")
        # fill serials, change to blue, and set PASS
        for row_num in range(2, sn_size+2):
            # set the SN to start + displacement(row_num-2 since row start from 1 and row 1 is useless)
            gr_file_workbook.Sheets(2).Cells(
                row_num, 5).Value = sn_start+row_num-2
            # set to number, no digital
            gr_file_workbook.Sheets(2).Cells(
                row_num, 5).NumberFormat = "0"
            # set to blue
            gr_file_workbook.Sheets(2).Cells(
                row_num, 5).Font.Color = 16711680
            # put a PASS at the cordinating cell
            gr_file_workbook.Sheets(2).Cells(
                row_num, 12).Value = "PASS"
        # save the GR file
        gr_file_workbook.Save()
        gr_file.Quit()
        message(__name__, "GR FILE BUILT")


def build_GR() -> None:
    message(__name__, "START BUILDING GR FILE")
    GR_file_path = askopenfilename(title="Select file to attach")
    if GR_file_path == "":
        return
    gr_file_base_dir = GR_file_path.replace(
        path.basename(GR_file_path), "")
    gr_file = win32com.client.Dispatch("Excel.Application")
    gr_file_workbook = gr_file.Workbooks.Open(GR_file_path)
    # for easier access
    gr_file_sheet1 = gr_file_workbook.Sheets(1)
    gr_file_sheet2 = gr_file_workbook.Sheets(2)
    gr_file_sheet3 = gr_file_workbook.Sheets(3)
    # the PB Number of this GR
    GR_PB = gr_file_sheet2.Cells(2, 1).Value
    # now the GR file is ready, need to file data
    sn_start = sn_end = 0
    # get range of the sn (because this may also be a GR file, which does not knowns range)
    row_number = 2
    max_columns = gr_file_sheet2.UsedRange.Columns.Count
    sn_col = 5
    for col in range(1, max_columns + 1):
        cell_value = str(gr_file_sheet2.Cells(row_number, col).Value)
        if re.match(SERIAL_REGEX, cell_value) != None:
            sn_start = int(float(cell_value))
            sn_end = sn_start+int(gr_file_sheet2.Cells(2, 4).Value)
            sn_col = col
    print(sn_start, sn_end)
    # remove contents of the sheet 3
    gr_file_sheet3.Cells.Clear()
    # put SN and STT at sheet 3
    gr_file_sheet3.Cells(1, 1).Value = "SN"
    gr_file_sheet3.Cells(1, 2).Value = "STT"
    message(__name__, "START SCANNING")
    # now scan the SN of the boards
    message(__name__, "PLEASE START SCANNING SN.ENTER \"quit\" TO QUIT. WHEN FINISH, PRESS ENTER TO CONTINUE")
    sn_set = set()
    while True:
        sn = input()
        try:
            if sn.lower() == "quit":
                message(__name__, "SCANNING STOPPED BY USER")
                return
            # end of scanning
            if sn == "":
                # check number of boards scanned
                message(
                    __name__, "{} UNITS, CORRECT? (PRESS ENTER TO CONTINUE, quit to Quit, ANYTHING ELSE TO SCAN MORE)".format(len(sn_set)))
                # scan more
                if input() == "":
                    break
                # no, need to scan more since something was wrong
                else:
                    continue
            # SN not in this wo range
            if int(sn) > sn_end or int(sn) < sn_start:
                alert_beep("SN NOT IN RANGE, PLEASE RESCAN LAST BOARD")
                continue
            # valid SN, add to set(this can avoid rescan same sn twice)
            sn_set.add(int(sn))
        # sn cannot be convert to int, happens when scan to the small label which contains characters
        except:
            alert_beep("SN NOT VALID, PLEASE RESCAN LAST BOARD")
            continue
    message(__name__, "SCANNING COMPLETE, START PROCESSING DATA")
    # get the sorted SN
    sn_list = sorted(sn_set)
    # fill the sheet 3
    row_num = 2
    for sn in sn_list:
        gr_file_sheet3.Cells(row_num, 1).Value = sn
        gr_file_sheet3.Cells(row_num, 1).NumberFormat = "0"
        gr_file_sheet3.Cells(row_num, 2).Value = "PASS"
        row_num += 1
    # remove sheet1, copy title from sheet2 to 1
    gr_file_sheet1.Cells.Clear()
    sn_col = gr_file_sheet2.Columns(sn_col)
    gr_file_sheet2.Rows(1).Copy()
    gr_file_sheet1.Rows(1).PasteSpecial(Paste=-4163)
    # show all data for searching
    try:
        gr_file_sheet2.ShowAllData()
    except:
        pass
    # start from row 2
    sheet1_row_counter = 2
    for sn in sn_list:
        res = sn_col.Find(
            What=sn,
            LookIn=-4163,
            LookAt=1,
            SearchOrder=1,
            SearchDirection=1
        )
        if res == None:
            message(__name__, "SN NOT FOUND")
            continue
        # get the row, change SN to red
        gr_file_sheet2.Cells(res.Row, 5).Font.Color = 255
        gr_file_sheet2.Rows(res.Row).Copy()
        gr_file_sheet1.Rows(sheet1_row_counter).PasteSpecial(Paste=-4163)
        gr_file_sheet1.Rows(sheet1_row_counter).NumberFormat = "0"
        sheet1_row_counter += 1
    message(__name__, "PROCESSING COMPLETE, FINALIZING GR FILE")
    formula_cell = gr_file_sheet2.Rows(2).Find(
        What="=",            # Most formulas start with "="
        LookIn=-4123,        # xlFormulas
        LookAt=1,            # xlWhole
        SearchOrder=1,       # xlByRows
        SearchDirection=1    # xlNext
    )
    try:
        # delete COL ret if exist
        gr_file_sheet1.Columns(formula_cell.Column).Delete()
    except:
        pass
    now = datetime.now()
    today_date = now.strftime("%m")+now.strftime("%d")
    # file name to save
    final_GR_file_name = r"{}{} {}_{}x_{}.xlsx".format(gr_file_base_dir, today_date, GR_PB, str(
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
