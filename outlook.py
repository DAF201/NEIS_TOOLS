import win32com.client
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from os import path
from config import CONFIG, DN_REGEX
from helpers import message, alert, get_line
import re
Tk().withdraw()  # Hide the root window

# get outlook
message(__name__, "LOADING OUTLOOK")
OUTLOOK = win32com.client.Dispatch("Outlook.Application")
OUTLOOK_INBOX = OUTLOOK.GETNAMESPACE("MAPI").GetDefaultFolder(6).Items
message(__name__, "OUTLOOK LOADED")


def get_unread_mails() -> list:
    """get the latest emails, 20 by default"""
    message(__name__, "READING EMAILS")
    # results buffer
    results = []
    # sort the mails by recv time
    OUTLOOK_INBOX.Sort("[ReceivedTime]", True)
    # get last 20 mails
    mails = [OUTLOOK_INBOX[i]
             for i in range(CONFIG["Outlook"]["NUM_OF_EMAILS_TO_READ"])]
    index = 0
    for mail in mails:
        # Only MailItem (class 43), skip calendar etc.
        if mail.Class != 43:
            continue
        # how many mails you want
        if index > CONFIG["Outlook"]["NUM_OF_EMAILS_TO_READ"]:
            break
        results.append({"sender": mail.SenderName, "address": mail.SenderEmailAddress,
                       "subject": mail.Subject, "content": mail.Body[:200]})
        index += 1
    message(__name__, "READING COMPLETE")
    return results


# send Email to request for ITN

def request_for_ITN() -> bool:
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

# send email to get a DN for a GR, all data comes from the excel


def request_for_DN() -> str:
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
            file_data[1], file_data[0], file_data[2].replace(".xlsx", ""), get_GR_status(attachment_path), CONFIG["Outlook"]["USER"])
        # add attachment
        mail.Attachments.Add(attachment_path)
        # send
        mail.Send()
        # return name
        return file_data[0]
    except:
        return ""


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
            {"".join(rows)}
        </table>
        </body>
    </html>
    """
