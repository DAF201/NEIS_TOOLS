import win32com.client
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from os import path
from excel import get_GR_status
from config import CONFIG, DN_REGEX
from print_log import message, alert, get_line
import re
Tk().withdraw()  # Hide the root window

# get outlook
message(__name__, "LOADING OUTLOOK")
OUTLOOK = win32com.client.Dispatch("Outlook.Application")
OUTLOOK_INBOX = OUTLOOK.GETNAMESPACE("MAPI").GetDefaultFolder(6).Items
message(__name__, "OUTLOOK LOADED")

# get the lastest mails, for check DN use


def get_unread_mails() -> list:

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
                       "subject": mail.Subject, "content": mail.Body[:500]})
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
        file_name = file_name.split(' ')[-1]
        # PB-61258_179x_5202A50110
        file_data = file_name.split('_')
        # ["PB-61258","179x","5202A50110"]

        # create a new email
        mail = OUTLOOK.CreateItem(0)

        # who to send to
        mail.To = CONFIG["DN"]["DN_mail_To"]
        mail.CC = CONFIG["DN"]["DN_mail_CC"]

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
