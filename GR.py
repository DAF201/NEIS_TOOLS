import win32com.client
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from config import re, SERIAL_REGEX
from print_log import *
Tk().withdraw()


def build_GR():
    GR_file_path = askopenfilename(title="Select file to attach")
    if GR_file_path == "":
        return
    GR_FILE = win32com.client.Dispatch("Excel.Application")
    GR_FILE_WORKBOOK = GR_FILE.Workbooks.Open(GR_file_path)

    # for easier access
    GR_FILE_SHEET1 = GR_FILE_WORKBOOK.Sheets(1)
    GR_FILE_SHEET2 = GR_FILE_WORKBOOK.Sheets(2)
    GR_FILE_SHEET3 = GR_FILE_WORKBOOK.Sheets(3)

    # egt the starting sn number index
    sn_start = 0
    # get size of excel
    sn_size = int(GR_FILE_SHEET2.Cells(2, 4))
    # sn end, which is the last sn of this wo, all sn should be with in, let say 100 units, start 01, the max is 100, so 01+100-1 is the max sn
    sn_end = sn_start+sn_size-1

    if GR_FILE_WORKBOOK.Sheets.Count == 1:
        # this file is a feed file, and need to create other 2 sheets

        # create gr copy sheet and SN STT sheet
        GR_FILE_WORKBOOK.Sheets.Add(
            Before=GR_FILE_SHEET1).Name = "Sheet 2"

        GR_FILE_WORKBOOK.Sheets.Add(
            After=GR_FILE_SHEET2).Name = "Sheet 1"

        # move the original sheet to index 2
        GR_FILE_SHEET3.Move(Before=GR_FILE_SHEET2)

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

        # put SN and STT at sheet 3
        GR_FILE_SHEET3.Cells(1, 1).Value = "SN"
        GR_FILE_SHEET3.Cells(1, 2).Value = "STT"

        # save the GR file
        GR_FILE_WORKBOOK.Save()

    # now the GR file is ready, need to file data

    # get number of rows has data in sheet 3
    used_range = GR_FILE_SHEET3.UsedRange
    last_row = used_range.Row + used_range.Rows.Count - 1  # Actual last used row
    # remove them
    GR_FILE_SHEET3.Range(
        GR_FILE_SHEET3.Cells(2, 1),
        GR_FILE_SHEET3.Cells(last_row, 2)
    ).ClearContents()

    # now scan the SN of the boards
    message(__name__, "PLEASE START SCANNING SN. WHEN FINISH, PRESS ENTER TO CONTINUE")
    sn_set = set()
    while True:
        sn = input()
        try:
            # end of scanning
            if sn == "":
                # check number of boards scanned
                message(
                    __name__, "{} UNITS, CORRECT? PRESS ENTER TO CONTINUE, ENTER ANYTHING ELSE TO SCAN MORE".format(len(sn_set)))
                # comfirm number, continue to process
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

    # get the sorted SN
    sn_list = sorted(sn_set)

    # fill the sheet 3
    row_num = 2
    for sn in sn_list:
        GR_FILE_SHEET3.Cells(row_num, 1).Value = sn
        GR_FILE_SHEET3.Cells(row_num, 1).NumberFormat = "0"
        GR_FILE_SHEET3.Cells(row_num, 2).Value = "PASS"
        row_num += 1

    # TODO if a sn exist in sheet2 and 3, copy the row in sheet2 to sheet1 next free row

    GR_FILE_WORKBOOK.Save()
