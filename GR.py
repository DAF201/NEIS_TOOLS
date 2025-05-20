import win32com.client
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from config import re, SERIAL_REGEX, CONFIG, update_config
from print_log import *
from os import path, rename
from datetime import datetime
Tk().withdraw()


# I know this function's variables naming were garbage but this function is too long, and the logic requires many var with similar names
def build_GR() -> None:
    message(__name__, "START BUILDING GR FILE")
    GR_file_path = askopenfilename(title="Select file to attach")
    if GR_file_path == "":
        return

    GR_FILE_BASE_DIRECTORY = GR_file_path.replace(
        path.basename(GR_file_path), '')

    GR_FILE = win32com.client.Dispatch("Excel.Application")
    GR_FILE_WORKBOOK = GR_FILE.Workbooks.Open(GR_file_path)

    # for easier access
    GR_FILE_SHEET1 = GR_FILE_WORKBOOK.Sheets(1)
    GR_FILE_SHEET2 = GR_FILE_WORKBOOK.Sheets(2)
    GR_FILE_SHEET3 = GR_FILE_WORKBOOK.Sheets(3)

    # the PB Number of this GR
    GR_PB = GR_FILE_SHEET2.Cells(2, 1).Value

    # get the starting sn number index, place holder
    sn_start = 0

    # get range of sn
    sn_size = int(GR_FILE_SHEET2.Cells(2, 4))

    # sn end, which is the last sn of this wo, all sn should be with in, let say 100 units, start 01, the max is 100, so 01+100-1 is the max sn
    sn_end = sn_start+sn_size-1

    # if this file is a feed file, then need to create other 2 sheets
    if GR_FILE_WORKBOOK.Sheets.Count == 1:

        message(__name__, "FEEDFILE DETECTED, START BUILDING GR FILE FROM FEEDFILE")

        # create gr copy sheet and SN STT sheet
        GR_FILE_WORKBOOK.Sheets.Add(
            Before=GR_FILE_SHEET1).Name = "Sheet 2"

        GR_FILE_WORKBOOK.Sheets.Add(
            After=GR_FILE_SHEET2).Name = "Sheet 1"

        # move the original sheet to index 2
        GR_FILE_SHEET3.Move(Before=GR_FILE_SHEET2)

        message(__name__, "COLLECTING INFOMATION")
        # get SN start until get a valid sn or quit command
        while (True):
            sn_start = input(
                "please enter strating serial number, enter quit to Quit")

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
            GR_FILE_SHEET2.Cells(
                row_num, 5).Value = sn_start+row_num-2

            # set to number, no digital
            GR_FILE_SHEET2.Cells(
                row_num, 5).NumberFormat = "0"

            # set to blue
            GR_FILE_SHEET2.Cells(
                row_num, 5).Font.Color = 16711680

            # put a PASS at the cordinating cell
            GR_FILE_SHEET2.Cells(
                row_num, 12).Value = "PASS"

        # save the GR file
        GR_FILE_WORKBOOK.Save()
        message(__name__, "GR FILE BUILT")

    # now the GR file is ready, need to file data

    # get range of the sn (because this may also be a GR file, which does not knowns range)
    sn_start = int(GR_FILE_SHEET2.Cells(2, 5).Value)
    sn_end = sn_start+int(GR_FILE_SHEET2.Cells(2, 4).Value)-1

    # remove the sheet 3
    GR_FILE_SHEET3.Cells.Clear()

    # put SN and STT at sheet 3
    GR_FILE_SHEET3.Cells(1, 1).Value = "SN"
    GR_FILE_SHEET3.Cells(1, 2).Value = "STT"

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
        GR_FILE_SHEET3.Cells(row_num, 1).Value = sn
        GR_FILE_SHEET3.Cells(row_num, 1).NumberFormat = "0"
        GR_FILE_SHEET3.Cells(row_num, 2).Value = "PASS"
        row_num += 1

    # remove sheet1, copy title from sheet2 to 1
    GR_FILE_SHEET1.Cells.Clear()
    sn_col = GR_FILE_SHEET2.Columns(5)

    GR_FILE_SHEET2.Rows(1).Copy()
    GR_FILE_SHEET1.Rows(1).PasteSpecial(Paste=-4163)

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

        # get the row, change SN to red
        GR_FILE_SHEET2.Cells(res.Row, 5).Font.Color = 255
        GR_FILE_SHEET2.Rows(res.Row).Copy()
        GR_FILE_SHEET1.Rows(sheet1_row_counter).PasteSpecial(Paste=-4163)
        GR_FILE_SHEET1.Rows(sheet1_row_counter).NumberFormat = "0"
        sheet1_row_counter += 1

    message(__name__, "PROCESSING COMPLETE, FINALIZING GR FILE")

    # delete COL ret if exist
    GR_FILE_SHEET1.Columns(16).Delete()

    now = datetime.now()
    today_date = now.strftime("%m")+now.strftime("%d")
    # file name to save
    final_GR_file_name = r"{}{} {}_{}x_{}.xlsx".format(GR_FILE_BASE_DIRECTORY, today_date, GR_PB, str(
        len(sn_list)), CONFIG["GR"]["invoice_header"]+str(CONFIG["GR"]["invoice_record"]).rjust(4, '0'))

    # save data
    GR_FILE_WORKBOOK.Save()
    GR_FILE_WORKBOOK.Close(SaveChanges=False)
    GR_FILE.Quit()

    # save next invoice
    CONFIG["GR"]["invoice_record"] = CONFIG["GR"]["invoice_record"]+1
    update_config()

    # rename file
    rename(GR_file_path, final_GR_file_name)

    message(__name__, "GR FILE BUILDING COMPLETE, FIEL CREATE: {}".format(
        final_GR_file_name))
