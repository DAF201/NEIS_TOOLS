import win32com.client
from gc import collect
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
TARGET_TABLE_workbook = TARGET_TABLE.Workbooks.Open(
    CONFIG["Excel"]["target_table"])
TARGET_TABLE_SHEET = TARGET_TABLE_workbook.Sheets(1)

BUFFER_TABLE = win32com.client.Dispatch("Excel.Application")
try:
    BUFFER_TABLE.Visible = False
except:
    pass
BUFFER_TABLE_workbook = BUFFER_TABLE.Workbooks.Open(
    CONFIG["Excel"]["buffer_table"])
BUFFER_TABLE_SHEET = BUFFER_TABLE_workbook.Sheets(1)

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

    # read column A and B
    while True:
        a_value = sheet.Cells(row, 1).Value
        b_value = sheet.Cells(row, 2).Value
        # Stop when column A is empty
        if a_value is None:
            break
        # Convert A to string to preserve formatting
        GR_data.append(((str(a_value))[:-2], b_value))
        row += 1

    # Close excel workbook without saving changes
    workbook.Close(SaveChanges=False)
    # close excel to release meoery
    excel.Quit()

    message(__name__, "READING COMPLETE, {} SN READ".format(len(GR_data)))

    # clean grabage
    collect()

    return build_html_table(GR_data)


def build_html_table(data: list) -> str:

    message(__name__, "STARTING BUILDING EMBEDDED HTML TABLE")
    # Table styles: border for all cells, center align
    table_style = "border-collapse: collapse;"
    cell_style = "border: 1px solid black; padding: 4px; text-align: center;"

    # Create table header
    header = f"<tr><th style='{cell_style}'>SN</th><th style='{cell_style}'>Status</th></tr>"

    # Create table rows from data
    rows = []
    for a, b in data:
        # Ensure column A is treated as a string
        rows.append(
            f"<tr><td style='{cell_style}'>{a}</td><td style='{cell_style}'>{b}</td></tr>")

    message(__name__, "BUILDING COMPLETE")

    # Return the full HTML table
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


def search_excel_PB(transaction: str | dict) -> int:
    # this will search for the last appearance of the transaction based on PB number

    # if not dictionary, load as dictionary
    message(__name__, "SEARCHING FOR TRANSACTION: {}".format(
        transaction['PB']))
    if type(transaction) == str:
        transaction = loads(transaction)

    column_range = TARGET_TABLE_SHEET.Columns(3)

    found = column_range.Find(
        What=transaction["PB"],
        LookIn=1,
        LookAt=1,
        SearchDirection=-4121
    )

    if found:
        message(__name__, "TRANSACTION FOUND AT ROW: {}".format(found.Row))
        return found.Row
    else:
        message(__name__, "TRANSACTION NOT FOUND")
        return 0


def search_excel_DN(transaction: str | dict) -> bool:
    if type(transaction) == str:
        transaction = loads(transaction)
    try:
        search_range = BUFFER_TABLE_SHEET.Range("L:L")
        found = search_range.Find(What=transaction["DN"], LookIn=1)
    except:
        return True
    return found == None


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
    BUFFER_TABLE_workbook.Save()


def edit_buffer_table(transaction: str | dict, row_number: int) -> int:
    # this will search for the last appearance of the transaction based on PB number

    # if not dictionary, load as dictionary
    if type(transaction) == str:
        transaction = loads(transaction)

    BUFFER_TABLE_SHEET.Cells(row_number, 12).Value = transaction["DN"]
    BUFFER_TABLE_SHEET.Cells(row_number, 13).Value = transaction["NUM"]
    BUFFER_TABLE_SHEET.Cells(row_number, 14).Value = transaction["DEST"]

    # Save changes
    BUFFER_TABLE_workbook.Save()
