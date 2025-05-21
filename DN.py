import openai
from json import loads
from config import CONFIG
from outlook import get_unread_mails
from excel import GR_invoice_search, find_next_blank_row, copy_row, edit_buffer_table, clean_excel
from print_log import message, alert, get_line

openai.api_key = CONFIG["AI"]["openai_key"]

# provide email to AI, let the AI extract infomation


def get_DN_info(msg) -> dict:
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[
            {"role": "system", "content":
             """
             YOU ARE A HELPFUL AGENT AND YOU WILL HELP READING THE EAMILS.
             IN THE OUTPUT, YOU WILL OUTPUT EXACTLY LIKE \{"DN":********, "DEST":"**", PB:"PB-*****", "NUM":"***", "GR_NUMBER":"5202A*****", "VALID":"*"\}
             YOU WILL NEED TO REACH THROUGH THE EMAIL TO FIND THOSE INFORMATIONS AND FILL IN THE CORRECT PLACE.
             DN STAND FOR DELIVERY NUMBER, WHICH IS A 8 DIGITS NUMBER. SOMETIMES THE DN MAY CONCAT WITH DEST SUCH AS "87013010-SC". ALSO IF THERE IS ANY LEADING 0 IN DN, YOU NEED TO REMOVE THE LEADING 0.
             DEST STAND FOR DESTINATION, WHICH IS THE SHIPPING DESTINGATION, AND FOR CHINA, USE CN, FOR TAIWAN, USE TW, FOR HONGKONG, USE HK, IN THE VALUE.
             PB STAND FOR PB-NUMBER, WHICH IS A IDENTIFIER FOR COMPOENTS BEING SHIPPED, AND IT IS "PB-" CONCATE WITH A 5 DIGITS FOR NUMBER.
             NUM STAND FOR THE NUMBER OF THE COMPUNENTS TO BE SHIPPED, USUALLY REPRESENTED AS A DECIMAL NUMBER CONCAT WITH A CHARACTER 'x'. WHEN PROCESSING DATA, REMOVE THE "x".
             GR_NUMBER IS A UNIQUE ID FOR A TRANSACTION, AND IT ALWAYS START WITH 5202A THEN CONCAT TO 5 DIGITS OF NUMBERS.
             VALID STAND FOR IF YOU CAN FIND ALL VALUES FROM THE EMAIL. IF THERE IS AT LEAST ONE VALUE YOU CANNOT FIND, FILL WITH NA, AND SET VALID="0". NO MAKE UP DATA IF YOU CANNOT FIND THE COORDINATE DATA FROM EMAIL.
             IF YOU CAN FIND ALL VALUES FROM THE EMAIL, VALID="1".
             PLASE NOTE, FOR DESTINATION, THE ZANKER AND SC ARE THE SAME, YOU SHOULD USE SC FOR BOTH CASE.
             ALSO, WHEN YOU CANNOT FIND DN, OR TOLD NO DN, IF THE EMAIL SAID FXSJ, VENDER POOL, STOCK, THIS MEANS THE DN IS "FXSJ" AND DEST IS ALSO "FXSJ", AND THIS IS STILL VALID=1 CASE.
             IN ALL CASES, YOUR RETURN VALUES SHOULD BE STRING, AND THE VALUE SHOULD BE ROUNDED BY " TO MAKE JSON LOADABLE.
             """},
            {"role": "user", "content": msg}
        ],
        temperature=0
    )
    try:
        return loads(response["choices"][0]["message"]["content"])
    except:
        return {}


def check_new_DN():
    message(__name__, "CLEANING BUFFER TABLE")
    clean_excel()
    message(__name__, "ALL CLEAN")

    message(__name__, "READING DN")
    emails = get_unread_mails()
    dn_info = []
    valid_dn = []
    for email in emails:
        dn_info.append(get_DN_info(str(email)))
    for dn in dn_info:
        try:
            if dn["VALID"] == '1':
                message(__name__, dn)
                valid_dn.append(dn)
        except:
            continue
    message(__name__, "READING COMPLETE")
    return valid_dn


def append_new_DN_to_excel(new_dn: list):
    message(__name__, "APPENDING DN INFOMATION TO BUFFER TABLE")
    for dn in new_dn:
        try:
            # search for the PB, get all possbile rows
            source_row = GR_invoice_search(dn)
            if source_row == 0:
                continue
            # for each row, copy to buffer, and edit data
            # for source_row in source_rows:
            target_row = find_next_blank_row()
            copy_row(source_row, target_row)
            edit_buffer_table(dn, target_row)
        except Exception as e:
            alert(__name__, get_line(), e)
            continue
    message(__name__, "DN INFORMATION ADDED ")
