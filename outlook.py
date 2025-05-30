import win32com.client
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from os import path
from config import CONFIG, DN_REGEX
from helpers import message, alert, get_line
import re
Tk().withdraw()  # Hide the root window

# get outlook, since this will work on the same outlook object, just make it global
message(__name__, "LOADING OUTLOOK")
OUTLOOK = win32com.client.Dispatch("Outlook.Application")
OUTLOOK_INBOX = OUTLOOK.GETNAMESPACE("MAPI").GetDefaultFolder(6).Items
message(__name__, "OUTLOOK LOADED")


def get_unread_mails() -> list:
    """get the latest emails, 20 by default"""
    message(__name__, "READING EMAILS")

    results = []
    OUTLOOK_INBOX.Sort("[ReceivedTime]", True)
    mails = [OUTLOOK_INBOX[i]
             for i in range(CONFIG["Outlook"]["NUM_OF_EMAILS_TO_READ"])]
    index = 0
    for mail in mails:
        # Only MailItem (class 43), skip calendar etc.
        if mail.Class != 43:
            continue
        if index > CONFIG["Outlook"]["NUM_OF_EMAILS_TO_READ"]:
            break
        results.append({"sender": mail.SenderName, "address": mail.SenderEmailAddress,
                       "subject": mail.Subject, "content": mail.Body[:200]})
        index += 1
    message(__name__, "READING COMPLETE")
    return results


def request_for_ITN() -> bool:
    """send Email to ITN, Need to attach SLI"""
    try:
        message(__name__, "INPUT DN FOR APPLYING FOR ITN NUMBER")
        dn = ""
        while True:
            dn = input("\t\t\t")
            if dn == "" or dn.lower() == "quit":
                message(__name__, "ITN APPLICATION CANCELLED")
                return False
            if re.match(DN_REGEX, dn) != None:
                break
        attachment_path = askopenfilename(title="Select file to attach")
        mail = OUTLOOK.CreateItem(0)
        # mail.To = CONFIG["DN"]["DN_TEST_ADDRESS"]
        # mail.CC = CONFIG["DN"]["DN_TEST_ADDRESS"]
        mail.To = CONFIG["ITN"]["ITN_mail_To"]
        mail.CC = CONFIG["ITN"]["ITN_mail_CC"]
        mail.Subject = CONFIG["ITN"]["ITN_SUBJECT"].format(dn)
        mail.Body = CONFIG["ITN"]["ITN_BODY"].format(
            dn, CONFIG["Outlook"]["USER"])
        mail.Attachments.Add(attachment_path)
        mail.Send()
        return True
    except Exception as e:
        alert(__name__, get_line(), e)
        return False


def request_for_DN() -> str:
    """send email to get a DN"""
    try:
        # open a window to select a file
        attachment_path = askopenfilename(title="Select file to attach")
        # get file name to extract data
        file_name = path.basename(attachment_path)
        # 0513 (1) PB-61258_179x_5202A0110
        file_name = file_name.split(" ")[-1]
        # PB-61258_179x_5202A50110
        file_data = file_name.split("_")
        # ["PB-61258","179x","5202A50110"]

        # create a new email
        mail = OUTLOOK.CreateItem(0)
        # who to send to
        mail.To = CONFIG["DN"]["DN_mail_To"]
        mail.CC = CONFIG["DN"]["DN_mail_CC"]
        # mail.To = CONFIG["DN"]["DN_TEST_ADDRESS"]
        # mail.CC = CONFIG["DN"]["DN_TEST_ADDRESS"]

        # subject
        mail.Subject = CONFIG["DN"]["DN_SUBJECT"].format(
            file_data[1], file_data[0], file_data[2].replace(".xlsx", ""))

        # get HTML body from the config file, fill with data read from excel
        mail.HTMLBody = CONFIG["DN"]["DN_HTML_BODY"].format(
            file_data[1], file_data[0], file_data[2].replace(".xlsx", ""), get_SN_STT(attachment_path), CONFIG["Outlook"]["USER"])

        mail.Attachments.Add(attachment_path)
        mail.Send()
        # just to check if send success
        return file_data[0]
    except:
        return ""


def get_SN_STT(path: str) -> str:
    """get the SN STT sheet and attach to email"""

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

    # Read A and B columns until A has no GR_data (sometime we have more pass than SN)
    row = 2
    GR_data = []

    # sometimes the sheet has config, then need to read 3 cols (Do those 4 cols sheet manually)
    has_config = (sheet.Cells(1, 3).Value != None)

    # read column A and B
    while True:
        a_value = sheet.Cells(row, 1).Value
        b_value = sheet.Cells(row, 2).Value
        c_value = ""

        if has_config:
            c_value = sheet.Cells(row, 3).Value

        # Stop when column A is empty
        if a_value is None:
            break

        # the SN alwawys attached with .00
        if has_config:
            GR_data.append(((str(a_value))[:-2], b_value, (str(c_value))[:-2]))
        else:
            GR_data.append(((str(a_value))[:-2], b_value))

        row += 1

    workbook.Close(False)
    excel.Quit()

    message(__name__, "READING COMPLETE, {} ROWS READ".format(len(GR_data)))

    # return HTML table based on if has config
    if not has_config:
        return build_html_table(GR_data)

    return build_html_table(GR_data, 1)


def build_html_table(data: list, has_config=0) -> str:
    """The email need to insert HTML to show a table"""

    message(__name__, "BUILDING SN STT TABLE")

    table_style = "border-collapse: collapse;"
    cell_style = "border: 1px solid black; padding: 4px; text-align: center;"

    rows = []

    if not has_config:
        header = f"<tr><th style='{cell_style}'>SN</th><th style='{cell_style}'>Status</th></tr>"
        for a, b in data:
            rows.append(
                f"<tr><td style='{cell_style}'>{a}</td><td style='{cell_style}'>{b}</td></tr>")

    else:
        header = f"<tr><th style='{cell_style}'>SN</th><th style='{cell_style}'>Status</th><th style='{cell_style}'>Config</th></tr>"
        for a, b, c in data:
            rows.append(
                f"<tr><td style='{cell_style}'>{a}</td><td style='{cell_style}'>{b}</td><td style='{cell_style}'>{c}</td></tr>")

    message(__name__, "TABLE BUILT")

    return f"""
    <html>
        <body>
        <table style="{table_style}">
            {header}
            {"".join(rows)}
        </table>
        </body>
    </html>
    """
