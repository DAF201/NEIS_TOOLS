import win32com.client
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from os import path
from config import CONFIG, DN_REGEX
from helpers import message, alert, get_line
import re
Tk().withdraw()  # Hide the root window


def get_unread_mails():
    """get the latest emails, 20 by default"""

    outlook = win32com.client.Dispatch("Outlook.Application")
    outlook_inbox = outlook.GETNAMESPACE("MAPI").GetDefaultFolder(6).Items
    message(__name__, "READING EMAILS")

    results = []
    outlook_inbox.Sort("[ReceivedTime]", True)
    mails = [outlook_inbox[i]
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
