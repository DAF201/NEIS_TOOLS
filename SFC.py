from bs4 import BeautifulSoup
import requests
from config import CONFIG


def mo_query(model_name: str, module_number="") -> list | dict:

    url = CONFIG["SFC"]["mo_query_api"]
    form = CONFIG["SFC"]["mo_query_form"]
    form["modelName"] = model_name

    html_content = requests.post(url=url, data=form).content
    soup = BeautifulSoup(html_content.decode('big5'), 'html.parser')
    table = soup.find('table')

    # Extract headers as list of strings
    headers = [th.get_text(strip=True)
               for th in table.find_all('td', class_='tableheader')]

    # Extract all rows except header row
    data = []
    for tr in table.find_all('tr')[1:]:  # skip header row
        cells = tr.find_all('td')
        if len(cells) != len(headers):
            # skip rows that don't match header count (optional)
            continue
        row_dict = {}
        for header, cell in zip(headers, cells):
            # Join all text inside cell (including <font> or other tags)
            text = ''.join(cell.stripped_strings)
            row_dict[header] = text
        data.append(row_dict)

    # if looking for a specific working order number
    if module_number != "":
        for content in data:
            if content["Mo_Number"] == module_number:
                return content

    return data
