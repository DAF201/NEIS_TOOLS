# NOT USED DUE TO CONSOLE CANNOT PASTE TO EXCEL PROPERLY

# import openai
# from json import loads
# from config import CONFIG
# from outlook import get_unread_mails
# import win32com.client
# from helpers import message, alert, get_line
# from datetime import datetime
# openai.api_key = CONFIG["AI"]["openai_key"]
# TARGET_TABLE = TARGET_TABLE_WORKBOOK = TARGET_TABLE_SHEET = ""


# def excel_clean_up():
#     message(__name__, "CLEAR UP START")
#     try:
#         TARGET_TABLE_WORKBOOK.Close(False)
#     except:
#         pass
#     try:
#         TARGET_TABLE.Quit()
#     except:
#         pass
#     message(__name__, "CLEAN UP FINISH")


# def excel_init():
#     global TARGET_TABLE, TARGET_TABLE_SHEET, TARGET_TABLE_WORKBOOK
#     message(__name__, "LOADING EXCEL")
#     TARGET_TABLE = win32com.client.Dispatch("Excel.Application")
#     try:
#         TARGET_TABLE.Visible = False
#     except:
#         pass
#     TARGET_TABLE_WORKBOOK = TARGET_TABLE.Workbooks.Open(
#         CONFIG["Excel"]["target_table"])
#     TARGET_TABLE_SHEET = TARGET_TABLE_WORKBOOK.Sheets(1)
#     message(__name__, "EXCEL LOADED")


# def get_DN_info(msg) -> dict:
#     response = openai.ChatCompletion.create(
#         model="gpt-3.5-turbo",
#         messages=[
#             {"role": "system", "content":
#              """
#              YOU ARE A HELPFUL AGENT AND YOU WILL HELP READING THE EAMILS.
#              IN THE OUTPUT, YOU WILL OUTPUT EXACTLY LIKE \{"DN":********, "DEST":"**", PB:"PB-*****", "NUM":"***", "GR_NUMBER":"5202A*****", "VALID":"*"\}
#              YOU WILL NEED TO REACH THROUGH THE EMAIL TO FIND THOSE INFORMATIONS AND FILL IN THE CORRECT PLACE.
#              DN STAND FOR DELIVERY NUMBER, WHICH IS A 8 DIGITS NUMBER. SOMETIMES THE DN MAY CONCAT WITH DEST SUCH AS "87013010-SC". ALSO IF THERE IS ANY LEADING 0 IN DN, YOU NEED TO REMOVE THE LEADING 0.
#              DEST STAND FOR DESTINATION, WHICH IS THE SHIPPING DESTINGATION, AND FOR CHINA, USE CN, FOR TAIWAN, USE TW, FOR HONGKONG, USE HK, IN THE VALUE.
#              PB STAND FOR PB-NUMBER, WHICH IS A IDENTIFIER FOR COMPOENTS BEING SHIPPED, AND IT IS "PB-" CONCATE WITH A 5 DIGITS FOR NUMBER.
#              NUM STAND FOR THE NUMBER OF THE COMPUNENTS TO BE SHIPPED, USUALLY REPRESENTED AS A DECIMAL NUMBER CONCAT WITH A CHARACTER "x". WHEN PROCESSING DATA, REMOVE THE "x".
#              GR_NUMBER IS A UNIQUE ID FOR A TRANSACTION, AND IT ALWAYS START WITH 5202A THEN CONCAT TO 5 DIGITS OF NUMBERS.
#              VALID STAND FOR IF YOU CAN FIND ALL VALUES FROM THE EMAIL. IF THERE IS AT LEAST ONE VALUE YOU CANNOT FIND, FILL WITH NA, AND SET VALID="0". NO MAKE UP DATA IF YOU CANNOT FIND THE COORDINATE DATA FROM EMAIL.
#              IF YOU CAN FIND ALL VALUES FROM THE EMAIL, VALID="1".
#              PLASE NOTE, FOR DESTINATION, THE ZANKER AND SC ARE THE SAME, YOU SHOULD USE SC FOR BOTH CASE.
#              ALSO, WHEN YOU CANNOT FIND DN, OR TOLD NO DN, IF THE EMAIL SAID FXSJ, VENDER POOL, STOCK, THIS MEANS THE DN IS "FXSJ" AND DEST IS ALSO "FXSJ", AND THIS IS STILL VALID=1 CASE.
#              IN ALL CASES, YOUR RETURN VALUES SHOULD BE STRING, AND THE VALUE SHOULD BE ROUNDED BY " TO MAKE JSON LOADABLE.
#              """},
#             {"role": "user", "content": msg}
#         ],
#         temperature=0
#     )
#     try:
#         return loads(response["choices"][0]["message"]["content"])
#     except:
#         return {}


# def check_new_DN():
#     message(__name__, "READING DN")
#     emails = get_unread_mails()
#     dn_info = [get_DN_info(str(email)) for email in emails]
#     valid_dn = [dn for dn in dn_info if dn.get("VALID") == "1"]
#     message(__name__, "READING COMPLETE")
#     return valid_dn


# def invoice_search(transaction: str | dict) -> int:
#     message(__name__, "SEARCHING FOR: {}".format(
#         transaction["GR_NUMBER"]))
#     res = TARGET_TABLE_SHEET.Columns(22).Find(
#         What=transaction["GR_NUMBER"],
#         LookIn=-4163,
#         LookAt=1,
#         SearchOrder=1,
#         SearchDirection=1
#     )
#     if res == None:
#         message(__name__, "NO MATCH")
#         return 0
#     else:
#         message(__name__, "MATCH FOUND: {}".format(res.Row))
#         return res.Row


# def append_new_DN_to_excel(new_dn: list):
#     excel_init()
#     message(__name__, "APPENDING DN INFOMATION TO BUFFER TABLE")
#     for dn in new_dn:
#         try:
#             # search for the PB, get all possbile rows
#             source_row = invoice_search(dn)
#             if source_row == 0:
#                 continue
#             row_range = TARGET_TABLE_SHEET.Range(TARGET_TABLE_SHEET.Cells(
#                 source_row, 1), TARGET_TABLE_SHEET.Cells(source_row, 24))
#             values = row_range.Value
#             row_str = values[0][0].strftime(
#                 "%m/%d") if values[0][0] is not None else ""
#             col_index = 0
#             for cell in values[0]:
#                 if col_index != 0:
#                     row_str += "\t{}".format(str(cell))
#                 col_index += 1
#             print(row_str)
#         except Exception as e:
#             alert(__name__, get_line(), e)
#             continue
#     message(__name__, "DN INFORMATION CREATED ")
#     excel_clean_up()
