import win32com.client
from tkinter import Tk
from tkinter.filedialog import askopenfilename

Tk().withdraw()


def build_GR():
    GR_file_path = askopenfilename(title="Select file to attach")
    if GR_file_path == "":
        return
    GR_FILE = win32com.client.Dispatch("Excel.Application")
    TARGET_TABLE_WORKBOOK = GR_FILE.Workbooks.Open(GR_file_path)

    if TARGET_TABLE_WORKBOOK.Sheets.Count == 1:
        # is feed file
        # create gr copy sheet and SN STT sheet
        TARGET_TABLE_WORKBOOK.Sheets.Add(
            Before=TARGET_TABLE_WORKBOOK.Sheets(1))
        TARGET_TABLE_WORKBOOK.Sheets.Add(
            Before=TARGET_TABLE_WORKBOOK.Sheets(3))
        
        serial_start=input("please enter strating serial number")
    else:
        # is not feed file
        pass
