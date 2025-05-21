import win32com.client
from config import CONFIG
from json import loads
from print_log import message

# open excel target table

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


def get_GR_status(path: str) -> str:
    message(__name__, "LOADING GR FILE")

    # get excel
    excel = win32com.client.Dispatch("excel.Application")
    try:
        excel.Visible = False
    except:
        pass

    # open the selected excel from request DN
    workbook = excel.Workbooks.Open(path)

    # Get the last sheet
    sheet = workbook.Sheets(workbook.Sheets.Count)
    message(__name__, "GR FILE LOADED")
    message(__name__, "READING SN")

    # Read A and B columns until A has no GR_data(sometime we have more pass than SN)
    row = 2
    GR_data = []
    has_config = (sheet.Cells(1, 3).Value != None)

    # read column A and B
    while True:
        a_value = sheet.Cells(row, 1).Value
        b_value = sheet.Cells(row, 2).Value
        c_value = ""
        if sheet.Cells(row, 3).Value is not None:
            c_value = sheet.Cells(row, 3).Value

        # Stop when column A is empty
        if a_value is None:
            break

        # Convert A to string to preserve formatting
        if c_value == "":
            GR_data.append(((str(a_value))[:-2], b_value))
        else:
            GR_data.append(((str(a_value))[:-2], b_value, (str(c_value))[:-2]))
        row += 1

    # Close excel workbook without saving changes
    workbook.Close(SaveChanges=False)

    # close excel to release memory
    excel.Quit()
    message(__name__, "READING COMPLETE, {} ROWS READ".format(len(GR_data)))
    if not has_config:
        return build_html_table(GR_data)
    return build_html_table(GR_data, 1)


def build_html_table(data: list, has_config=0) -> str:

    message(__name__, "BUILDING SN STT TABLE")

    table_style = "border-collapse: collapse;"
    cell_style = "border: 1px solid black; padding: 4px; text-align: center;"
    rows = []
    if not has_config:

        header = f"<tr><th style='{cell_style}'>SN</th><th style='{cell_style}'>Status</th></tr>"

        for a, b in data:

            rows.append(
                f"<tr><td style='{cell_style}'>{a}</td><td style='{cell_style}'>{b}</td></tr>")

        message(__name__, "TABLE BUILT")
    else:
        header = f"<tr><th style='{cell_style}'>SN</th><th style='{cell_style}'>Status</th><th style='{cell_style}'>Config</th></tr>"

        for a, b, c in data:

            rows.append(
                f"<tr><td style='{cell_style}'>{a}</td><td style='{cell_style}'>{b}</td><td style='{cell_style}'>{c}</td></tr>")

    return f"""
    <html>
        <body>
        <table style="{table_style}">
            {header}
            {''.join(rows)}
        </table>
        </body>
    </html>
    """


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


def GR_invoice_search(transaction: str | dict) -> int:
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

    # Manually copy each cell's value
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


def clean_excel():
    try:
        BUFFER_TABLE_SHEET.Cells.Clear()
        BUFFER_TABLE_WORKBOOK.Save()
    except:
        pass

def excel_atexit_clean_up():
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
