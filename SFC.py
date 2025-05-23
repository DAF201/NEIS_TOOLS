from bs4 import BeautifulSoup
import requests
from config import CONFIG


def mo_query(model_name: str, module_number="") -> list:
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
