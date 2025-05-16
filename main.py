from DN import *
from test_msg import *
from outlook import *
from excel import *
from print_log import *


while (1):
    message(__name__, "Please select a step to continue:\n\t\t\t1. Sending DN request for a GR\n\t\t\t2. Checking Email for new DN\n\t\t\tEnter quit to Quit")
    process = input()
    if process == '1':
        prcessed_PB = request_for_DN()
        if prcessed_PB == "":
            message(__name__, "Email not sent due to no file being selected")
        else:
            message(__name__, "Request for {} sent".format(prcessed_PB))
    if process == '2':
        append_new_DN_to_excel(check_new_DN())
        message(__name__, "DN update complete")
    if process.lower() == 'quit':
        message(__name__, "Program Closing")
        break


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
