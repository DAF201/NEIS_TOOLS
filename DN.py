import openai
from json import loads
from config import CONFIG
from outlook import get_unread_mails
from excel import search_excel_PB, find_next_blank_row, copy_row, edit_buffer_table
from print_log import message
openai.api_key = CONFIG["AI"]["openai_key"]

# provide email to AI, let the AI extract infomation


def get_DN_info(msg) -> dict:
    response = openai.ChatCompletion.create(
        model="gpt-4-turbo",
        messages=[
            {"role": "system", "content":
             """
             YOU ARE A HELPFUL AGENT AND YOU WILL HELP READING THE EAMILS.
             IN THE OUTPUT, YOU WILL OUTPUT EXACTLY LIKE \{"DN":********, "DEST":"**", PB:"PB-*****", "NUM":"***", "VALID":"*"\}
             YOU WILL NEED TO REACH THROUGH THE EMAIL TO FIND THOSE INFORMATIONS AND FILL IN THE CORRECT PLACE.
             DN STAND FOR DELIVERY NUMBER, WHICH IS A 8 DIGITS NUMBER
             DEST STAND FOR DESTINATION, WHICH IS THE SHIPPING DESTINGATION
             PB STAND FOR PB-NUMBER, WHICH IS A IDENTIFIER FOR COMPOENTS BEING SHIPPED, AND IT IS "PB-" CONCATE WITH A 5 DIGITS FOR NUMBER.
             NUM STAND FOR THE NUMBER OF THE COMPUNENTS TO BE SHIPPED.
             VALID STAND FOR IF YOU CAN FIND ALL VALUES FROM THE EMAIL. IF THERE IS AT LEAST ONE VALUE YOU CANNOT FIND, FILL WITH NA, AND SET VALID="0".
             IF YOU CAN FIND ALL VALUES FROM THE EMAIL, VALID="1".
             PLASE NOTE, FOR DESTINATION, THE ZANKER AND SC ARE THE SAME, YOU SHOULD USE SC FOR BOTH CASE.
             ALSO, WHEN YOU CANNOT FIND DN, OR TOLD NO DN, IF THE EMAIL SAID FXSJ, VENDER POOL, STOCK, THIS MEANS THE DN AND DEST ARE BOTH "FXSJ", AND THIS IS STILL VALID=1 CASE.
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
    message(__name__, "FILTERING DN INFOMATION")
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
    message(__name__, "FILTERING COMPLETE")
    return valid_dn
    # return dn_info


def append_new_DN_to_excel(new_dn: list):
    message(__name__, "ADDING DN TO SHEET")
    for dn in new_dn:
        try:
            src_row = search_excel_PB(dn)
            target_row = find_next_blank_row()
            copy_row(src_row, target_row)
            edit_buffer_table(dn, target_row)
        except:
            continue
    message(__name__,"ADDING COMPLETE")
