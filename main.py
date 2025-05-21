from DN import *
from outlook import *
from excel import *
from print_log import *
from GR import *
import atexit

atexit.register(excel_atexit_clean_up)

while (1):
    message("main", "Please select a step to continue:\n\t\t\t"
            "1. Send DN request for a GR\n\t\t\t"
            "2. Check Email for new DN\n\t\t\t"
            "3. Build GR file\n\t\t\t"
            "4. Build GR file from FeedFile\n\t\t\t"
            "5. Apply ITN for a DN\n\t\t\t"
            "Enter quit to Quit")
    process = input("\t\t\t")
    match(process):
        case '1':
            if request_for_DN() == "":
                message(
                    "main", "Email not sent due to an exception happened while running")
            else:
                message("main", "Email sent")

        case '2':
            append_new_DN_to_excel(check_new_DN())
            message("main", "DN update complete")

        case '3':
            build_GR()
            message("main", "GR file ready")

        case '4':
            build_GR_from_feedfile()

        case '5':
            if request_for_ITN():
                message("main", "ITN request sent")
            else:
                message("main", "ITN request not sent")

        case 'quit':
            message("main", "Program closing")
            break

        case _:
            continue
