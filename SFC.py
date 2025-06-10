from bs4 import BeautifulSoup
import requests
from config import CONFIG


def mo_query(model_name: str, module_number="") -> list:
    """to look up the starting number of the board and make the GR file for the feedfile"""
    url = CONFIG["SFC"]["mo_query_api"]
    form = CONFIG["SFC"]["mo_query_form"]
    form["modelName"] = model_name
    html_content = requests.post(url=url, data=form).content
    soup = BeautifulSoup(html_content.decode("big5"), "html.parser")
    table = soup.find("table")
    headers = [th.get_text(strip=True)
               for th in table.find_all("td", class_="tableheader")]
    data = []
    for tr in table.find_all("tr")[1:]:
        cells = tr.find_all("td")
        if len(cells) != len(headers):
            continue
        row_dict = {}
        for header, cell in zip(headers, cells):
            text = "".join(cell.stripped_strings)
            row_dict[header] = text
        data.append(row_dict)
    if module_number != "":
        for content in data:
            if content["Mo_Number"] == module_number:
                return [content]
    return data


def WIP(working_order: str, department="OQC") -> list:
    """to lookup the information about a WO, for OQC it will return the cartoon id for scanning, for PACKING it will return the SN numbers for GR"""
    url = CONFIG["SFC"]["WIP_api"].format(department, working_order)
    html_content = requests.get(url).content.decode("big5", errors="ignore")
    soup = BeautifulSoup(html_content, "html.parser")
    rows = soup.find_all("tr")
    headers = [
        "ID", "Serial Number", "Mo Number", "Model Name", "Version Code",
        "Line Name", "Group Name", "Error Flag", "In Station Time",
        "Container NO", "Carton NO", "Emp Name"
    ]
    data = []
    for row in rows:
        cols = row.find_all("td")
        values = [col.get_text(strip=True).replace("\xa0", "") for col in cols]
        if len(values) == len(headers):
            row_dict = dict(zip(headers, values))
            data.append(row_dict)
    return data


def SN_look_up(sn):
    """For lookup information about a specific board"""
    if sn == "quit":
        return {}
    form = CONFIG["SFC"]["product_tracking"]["product_tracking_form"]
    form["T_SN"] = sn
    url = CONFIG["SFC"]["product_tracking"]["product_tracking_api"]
    html_content = requests.post(
        url, params=form).content.decode("big5", errors="ignore")
    soup = BeautifulSoup(html_content, "html.parser")
    rows = soup.find_all("tr")
    data = []
    for row in rows:
        cols = row.find_all("td")
        data.append([col.get_text(strip=True).replace("\xa0", "")
                    for col in cols])
    data = data[1:]
    # data belike:
    # ...
    # ['SN', "'1581925605005'", 'In_Line_Time', '2025/05/07 13:47:47']
    # ['MO_Number', '002100003109-1', 'Model_Name', '699-2G525-0220-TS5']
    # ...
    res = {"NPI OUT": False}
    i = 0
    for x in data:
        for i in range(0, len(x), 2):
            if i+1 < len(x):
                if "NPI_OUT" in x:
                    res["NPI OUT"] = True
                if "OQC" in x:
                    res["cartoon_id"] = x[8]
                res[x[i]] = x[i+1]
    return res
