import win32com.client
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from os import path
from excel import get_GR_status
from config import CONFIG
from print_log import message
Tk().withdraw()  # Hide the root window

# get outlook
message(__name__, "LOADING OUTLOOK")
OUTLOOK = win32com.client.Dispatch("Outlook.Application")
OUTLOOK_INBOX = OUTLOOK.GETNAMESPACE("MAPI").GetDefaultFolder(6).Items
message(__name__, "LOADING COMPLETE")

# get the lastest mails, for check DN use


def get_unread_mails() -> list:

    message(__name__, "STARTING WITH READING EMAILS")

    # results buffer
    results = []
    # sort the mails by recv time
    OUTLOOK_INBOX.Sort("[ReceivedTime]", True)
    # get last 20 mails
    mails = [OUTLOOK_INBOX[i]
             for i in range(CONFIG["DN"]["NUM_OF_EMAILS_TO_READ"])]
    index = 0
    for mail in mails:
        # Only MailItem (class 43), skip calendar etc.
        if mail.Class != 43:
            continue
        # how many mails you want
        if index > CONFIG["DN"]["NUM_OF_EMAILS_TO_READ"]:
            break
        results.append({"sender": mail.SenderName, "address": mail.SenderEmailAddress,
                       "subject": mail.Subject, "content": mail.Body[:500]})
        index += 1

    message(__name__, "READING COMPLETE")

    return results


# send email to get a DN for a GR, all data comes from the excel
def request_for_DN() -> str:
    try:
        # open a window to select a file
        attachment_path = askopenfilename(title="Select files to attach")
        # get file name to extract data
        file_name = path.basename(attachment_path)
        # 0513 (1) PB-61258_179x_5202A0110
        file_name = file_name.split(' ')[-1]
        # PB-61258_179x_5202A50110
        file_data = file_name.split('_')
        # ["PB-61258","179x","5202A50110"]

        # create a new email
        mail = OUTLOOK.CreateItem(0)

        # who to send to
        mail.To = CONFIG["DN"]["DN_TEST_ADDRESS"]
        mail.CC = CONFIG["DN"]["DN_TEST_ADDRESS"]

        # subject
        mail.Subject = "GR {} units of {} are done, please cut DN".format(
            file_data[1], file_data[0])

        # get HTML body from the config file, fill with data read from excel
        mail.HTMLBody = CONFIG["DN"]["DN_HTML_BODY"].format(
            file_data[1], file_data[0], get_GR_status(attachment_path), CONFIG["Outlook"]["USER"])

        # add attachment
        mail.Attachments.Add(attachment_path)

        # send
        mail.Send()

        # return name
        return file_data[0]
    except:
        return ""
