from DN import *
from outlook import *
from excel import *
from print_log import *
from GR import *
import atexit

atexit.register(excel_atexit_clean_up)

while (1):
    message("main", "Please select a step to continue:\n\t\t\t1. Sending DN request for a GR\n\t\t\t2. Checking Email for new DN\n\t\t\t3. Build GR file\n\t\t\tEnter quit to Quit")
    process = input("\t\t\t")
    match(process):
        case '1':
            if request_for_DN() == "":
                message(
                    "main", "Email not sent due to an exception happened while running")
            else:
                message("main", "Email for sent")

        case '2':
            append_new_DN_to_excel(check_new_DN())
            message("main", "DN update complete")

        case '3':
            build_GR()
            message("main", "GR file ready")

        case 'quit':
            message("main", "Program closing")
            break

        case _:
            continue
