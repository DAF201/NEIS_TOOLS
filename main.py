from DN import *
from outlook import *
from helpers import *
from GR import *
from SFC import *
import atexit

atexit.register(excel_atexit_clean_up)

while (1):
    message("main", "Please select a step to continue:\n\t\t\t"
            "1. Send DN request for a GR\n\t\t\t"
            "2. Check Email for new DN\n\t\t\t"
            "3. Build GR file\n\t\t\t"
            "4. Build GR file from FeedFile\n\t\t\t"
            "5. Look up Model information\n\t\t\t"
            "6. Carton number look up\n\t\t\t"
            "7. Apply ITN for a DN\n\t\t\t"
            "Enter quit to Quit")
    process = input("\t\t\t")
    match(process):
        case "1":
            if request_for_DN() == "":
                message(
                    "main", "Email not sent due to an exception happened while running")
            else:
                message("main", "Email sent")

        case "2":
            excel_init()
            append_new_DN_to_excel(check_new_DN())
            message("main", "DN update complete")

        case "3":
            build_GR()
            message("main", "GR file ready")

        case "4":
            build_GR_from_feedfile()

        case "5":
            message("main", "Please enter the model number")
            mo = input()
            message(
                "main", "Please enter the working order number, leave blank for all woking order number")
            wo = input()
            data = mo_query(mo, wo)
            message("main", "Searching result:")
            for line in data:
                message(
                    "main", "\n\t\t\tWorking Order : {}\t\t\tModel : {}\t\t\tTarget Quantity : {}\n\t\t\tSN Start : {}\t\t\tSN End : "
                    "{}\t\t\t\tCreate Date : {}".format(line["Mo_Number"], line["Model_Name"], line["Target_Qty"],
                                                        line["SN_Start"], line["SN_End"], line["Mo_Create_Date"]))
        case "6":
            message("main", "Please enter the working order number")
            data = carton_number(input())
            carton_nums = set()
            for line in data:
                carton_nums.add(line["Carton NO"])
                message(
                    "main", " Carton Number : {}\t\t  Containter Number : {}\t\tIn Stataion Time : {}".format(
                        line["Carton NO"], line["Container NO"],  line["In Station Time"]))
            carton_nums = sorted(carton_nums)
            message("main", "Below are Carton Numbers for CTRL CV \n")
            for carton_num in carton_nums:
                print(carton_num)
            print()
        case "7":
            if request_for_ITN():
                message("main", "ITN request sent")
            else:
                message("main", "ITN request not sent")

        case "quit":
            message("main", "Program closing")
            break

        case _:
            continue
