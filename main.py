from DN import *
from outlook import *
from helpers import *
from GR import *
from SFC import *
from config import *
from report_op import *
from pod import *
while (1):

    message("main", "Please select a step to continue:\n\t\t\t"
            "1. Send DN request for a GR\n\t\t\t"
            "2. Check Email for new DN (in progress, do not use) \n\t\t\t"
            "3. Build GR file\n\t\t\t"
            "4. Build GR file from Feedfile\n\t\t\t"
            "5. Model info look up\n\t\t\t"
            "6. Working order number look up\n\t\t\t"
            "7. Apply ITN for a DN\n\t\t\t"
            "8. Product Tracking by SN\n\t\t\t"
            "9. POD generate\n\t\t\t"
            "0. Create Report\n\t\t\t"
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
            append_new_DN_to_excel(check_new_DN())
            message("main", "DN update complete")

        case "3":
            build_GR()
            message("main", "GR file ready")

        case "4":
            build_GR_from_feedfile()
            message("main", "GR file built")

        case "5":
            message("main", "Please enter the model number")
            mo = ""
            while True:
                mo = input()
                if mo == "quit" or re.match(MO_REGEX, mo) != None:
                    break
                else:
                    message("main", "invalid MO")
            if mo == "quit":
                continue
            message("main", "Please enter the working order number, blank for all")
            wo = ""
            while True:
                wo = input()
                if wo == "" or wo == "quit" or re.match(WO_REFEX, wo) != None:
                    break
                else:
                    message("main", "invalid WO")
            if wo == "quit":
                continue
            data = mo_query(mo, wo)
            message("main", "Searching result:")
            for line in data:
                message(
                    "main", "\n\t\t\tWorking Order : {}\t\t\tModel : {}\t\t\tTarget Quantity : {}\n\t\t\tSN Start : {}\t\t\tSN End : "
                    "{}\t\t\t\tCreate Date : {}".format(line["Mo_Number"], line["Model_Name"], line["Target_Qty"],
                                                        line["SN_Start"], line["SN_End"], line["Mo_Create_Date"]))

        case "6":
            message(
                "main", "Please enter the department: OQC, PACKING. (More might be added in Future)")
            department = ""
            while True:
                department = input().upper()
                if department == "QUIT" or department in ["OQC", "PACKING"]:
                    break
                else:
                    message("main", "Invalid Department")
            if department == "QUIT":
                continue
            message("main", "Please enter the working order number")
            wo = ""
            while True:
                wo = input()
                if wo == "" or wo == "quit" or re.match(WO_REFEX, wo) != None:
                    break
                else:
                    message("main", "invalid WO")
            if wo == "quit":
                continue
            data = WIP(wo, department)
            if department == "OQC":
                carton_nums = set()
                for line in data:
                    carton_nums.add(line["Carton NO"])
                carton_nums = sorted(carton_nums)
                message("main", "Below are Carton Numbers for CTRL CV \n")
                for carton_num in carton_nums:
                    print(carton_num)
            elif department == "PACKING":
                message("main", "Below are Serial Numbers for CTRL CV \n")
                for line in data:
                    print(line["Serial Number"])
            print()

        case "7":
            if request_for_ITN():
                message("main", "ITN request sent")
            else:
                message("main", "ITN request not sent")

        case "8":
            message("main", "Please enter the working order number")
            res = SN_look_up(input("\t\t\t"))
            if res == {}:
                message("main", "cancelled")
            else:
                print("Serial Number:\t{}\t\t\tPBR:\t{}".format(
                    res["SN"][1:-1], res["PBR"]))
                print("Working Order Number:\t{}\t\t\tPart Number:\t{}".format(
                    res["MO_Number"], res["Model_Name"]))
                print(
                    "Next Station:\t{}\t\t\tNPI-OUT:\t{}".format(res["Next Station"], res["NPI OUT"]))
                if res.get("cartoon_id") != None:
                    print("Cartoon ID:\t{}".format(res["cartoon_id"]))

        case "9":
            message("main", "Downloading POD")
            make_pod()

        case "quit":
            message("main", "Program closing")
            break

        case "0":
            message("main", "Start create report")
            create_report()

        case _:
            continue
