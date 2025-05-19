import win32com.client
from config import CONFIG
from json import loads
from print_log import message

# open excel target table

message(__name__, "LOADING TARGET TABLE AND BUFFER TABLE")

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

message(__name__, "LOADING COMPLETE")


def get_GR_status(path: str) -> str:
    message(__name__, "LOADING GR TABLE")

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
    message(__name__, "LOADING COMPLETE")
    message(__name__, "READING GR SN")

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
    message(__name__, "READING COMPLETE, {} SN READ".format(len(GR_data)))
    if not has_config:
        return build_html_table(GR_data)
    return build_html_table(GR_data, 1)


def build_html_table(data: list, has_config=0) -> str:

    message(__name__, "STARTING BUILDING EMBEDDED HTML TABLE")

    table_style = "border-collapse: collapse;"
    cell_style = "border: 1px solid black; padding: 4px; text-align: center;"
    rows = []
    if not has_config:

        header = f"<tr><th style='{cell_style}'>SN</th><th style='{cell_style}'>Status</th></tr>"

        for a, b in data:

            rows.append(
                f"<tr><td style='{cell_style}'>{a}</td><td style='{cell_style}'>{b}</td></tr>")

        message(__name__, "BUILDING COMPLETE")
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
    message(__name__, "SEARCHING FOR TRANSACTION:{} IN TARGET TABLE".format(
        transaction["PB"]))

    if isinstance(transaction, str):
        transaction = loads(transaction)

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
            row, 10).Value)

        if num_in_row == NUM_value:
            message(__name__, "POSSIBLE MATCH FOUND AT ROW {}".format(row))
            row_nums.append(row)

        found = PB_col.FindNext(found)
        if found is None or found.Address == first_found.Address:
            break

    message(__name__, "SEARCHING DONE")

    return row_nums


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
    if transaction["DN"] == "SC":
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
    used_range = BUFFER_TABLE_SHEET.UsedRange
    last_row = used_range.Row + used_range.Rows.Count - 1  # Actual last used row
    last_col = used_range.Column + used_range.Columns.Count - 1  # Actual last used column

    BUFFER_TABLE_SHEET.Range(
        BUFFER_TABLE_SHEET.Cells(1, 1),
        BUFFER_TABLE_SHEET.Cells(last_row, last_col)
    ).ClearContents()
    BUFFER_TABLE_WORKBOOK.Save()
